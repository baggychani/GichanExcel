import {
  BooleanNumber,
  CellValueType,
  LocaleType,
  mergeWorksheetSnapshotWithDefault,
  type ICellData,
  type IRange,
  type IStyleData,
  type IWorkbookData,
} from "@univerjs/core";
import * as XLSX from "xlsx";
import { IMPORTED_LAYOUT_CUSTOM_KEY, UNIVER_MODEL_VERSION } from "./workbook-constants";
import {
  BUILT_IN_NUMBER_FORMATS,
  isStyleEmpty,
  readOpenXmlSheetInfo,
  readOpenXmlStyleInfo,
} from "./ooxml-styles";

export function workbookPreservesImportedLayout(workbookData: IWorkbookData): boolean {
  return workbookData.custom?.[IMPORTED_LAYOUT_CUSTOM_KEY] === true;
}

function sheetJsRangeToUniverRange(range: XLSX.Range): IRange {
  return {
    startRow: range.s.r,
    startColumn: range.s.c,
    endRow: range.e.r,
    endColumn: range.e.c,
  };
}

function sheetJsColumnToUniverColumn(column: XLSX.ColInfo) {
  return {
    ...(typeof column.wpx === "number"
      ? { w: column.wpx }
      : typeof column.width === "number"
        ? { w: Math.round(column.width * 6) }
        : {}),
    ...(column.hidden ? { hd: BooleanNumber.TRUE } : {}),
  };
}

function sheetJsRowToUniverRow(row: XLSX.RowInfo) {
  return {
    ...(typeof row.hpx === "number"
      ? { h: row.hpx, ia: BooleanNumber.FALSE }
      : typeof row.hpt === "number"
        ? { h: row.hpt, ia: BooleanNumber.FALSE }
        : {}),
    ...(row.hidden ? { hd: BooleanNumber.TRUE } : {}),
  };
}

function dateToExcelSerial(date: Date, date1904: boolean): number {
  const epoch = Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30);
  return (date.getTime() - epoch) / 86400000;
}

function sheetJsCellToUniverCell(cell: XLSX.CellObject, date1904: boolean): ICellData | null {
  const formula = cell.f ? (cell.f.startsWith("=") ? cell.f : `=${cell.f}`) : undefined;
  const cellFormat = cell.z && cell.z !== "General" ? String(cell.z) : "";
  const styleFromCellFormat: Pick<ICellData, "s"> =
    cellFormat ? { s: { n: { pattern: cellFormat } } } : {};

  if (cell.t === "n" && typeof cell.v === "number") {
    return {
      v: cell.v,
      t: CellValueType.NUMBER,
      ...styleFromCellFormat,
      ...(formula ? { f: formula } : {}),
    };
  }

  if (cell.t === "b" && typeof cell.v === "boolean") {
    return {
      v: cell.v,
      t: CellValueType.BOOLEAN,
      ...styleFromCellFormat,
      ...(formula ? { f: formula } : {}),
    };
  }

  if (cell.v instanceof Date) {
    return {
      v: dateToExcelSerial(cell.v, date1904),
      t: CellValueType.NUMBER,
      s: {
        n: {
          pattern: cellFormat || BUILT_IN_NUMBER_FORMATS[14],
        },
      },
      ...(formula ? { f: formula } : {}),
    };
  }

  const value = cell.v ?? cell.w ?? "";
  if (value === "" && !formula) {
    return null;
  }

  return {
    v: String(value),
    t: CellValueType.STRING,
    ...styleFromCellFormat,
    ...(formula ? { f: formula } : {}),
  };
}

function mergeCellStyles(...styles: Array<ICellData["s"] | undefined>): IStyleData | undefined {
  const merged: IStyleData = {};

  styles.forEach((style) => {
    if (!style || typeof style === "string") {
      return;
    }

    Object.assign(merged, style);
  });

  return isStyleEmpty(merged) ? undefined : merged;
}

function createStyleRegistry() {
  const styles: NonNullable<IWorkbookData["styles"]> = {};
  const styleIdsByKey = new Map<string, string>();

  return {
    styles,
    register(style: IStyleData | undefined): string | undefined {
      if (!style || isStyleEmpty(style)) {
        return undefined;
      }

      const key = JSON.stringify(style);
      const existing = styleIdsByKey.get(key);
      if (existing) {
        return existing;
      }

      const id = `import_style_${styleIdsByKey.size}`;
      styleIdsByKey.set(key, id);
      styles[id] = style;
      return id;
    },
  };
}

export async function importExcelWithSheetJs(file: File): Promise<IWorkbookData> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: false,
    cellNF: true,
    cellStyles: true,
  });
  const styleInfo = await readOpenXmlStyleInfo(buffer);
  const workbookId = crypto.randomUUID();
  const sheetOrder: string[] = [];
  const sheets: IWorkbookData["sheets"] = {};
  const styleRegistry = createStyleRegistry();
  const date1904 = Boolean(workbook.Workbook?.WBProps?.date1904);

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const openXmlSheetInfo = await readOpenXmlSheetInfo(styleInfo, sheetName);
    const sheetId = crypto.randomUUID();
    const cellData: IWorkbookData["sheets"][string]["cellData"] = {};
    const rowData: IWorkbookData["sheets"][string]["rowData"] = {};
    const columnData: IWorkbookData["sheets"][string]["columnData"] = {};
    const mergeData = (worksheet?.["!merges"] ?? []).map(sheetJsRangeToUniverRange);
    let maxRow = 0;
    let maxColumn = 0;

    if (worksheet?.["!ref"]) {
      const range = XLSX.utils.decode_range(worksheet["!ref"]);
      maxRow = range.e.r;
      maxColumn = range.e.c;
    }

    Object.keys(worksheet ?? {}).forEach((address) => {
      if (address.startsWith("!")) {
        return;
      }

      const { r: row, c: column } = XLSX.utils.decode_cell(address);
      const cell = sheetJsCellToUniverCell(worksheet[address] as XLSX.CellObject, date1904);
      if (!cell) {
        return;
      }

      const styleId = styleRegistry.register(
        mergeCellStyles(cell.s, openXmlSheetInfo?.cellStylesByRef.get(address)),
      );
      if (styleId) {
        cell.s = styleId;
      }

      cellData[row] ??= {};
      cellData[row][column] = cell;
      maxRow = Math.max(maxRow, row);
      maxColumn = Math.max(maxColumn, column);
    });

    worksheet?.["!rows"]?.forEach((row, rowIndex) => {
      if (!row) {
        return;
      }

      const rowInfo = sheetJsRowToUniverRow(row);
      if (Object.keys(rowInfo).length) {
        rowData[rowIndex] = rowInfo;
      }
    });

    worksheet?.["!cols"]?.forEach((column, columnIndex) => {
      if (!column) {
        return;
      }

      const columnInfo = sheetJsColumnToUniverColumn(column);
      if (Object.keys(columnInfo).length) {
        columnData[columnIndex] = columnInfo;
      }
    });

    sheetOrder.push(sheetId);
    sheets[sheetId] = mergeWorksheetSnapshotWithDefault({
      id: sheetId,
      name: sheetName,
      rowCount: Math.max(100, maxRow + 1),
      columnCount: Math.max(26, maxColumn + 1),
      mergeData,
      cellData,
      rowData,
      columnData,
    });
  }

  if (sheetOrder.length === 0) {
    const sheetId = crypto.randomUUID();
    sheetOrder.push(sheetId);
    sheets[sheetId] = mergeWorksheetSnapshotWithDefault({
      id: sheetId,
      name: "Sheet1",
      rowCount: 100,
      columnCount: 26,
      cellData: {},
    });
  }

  return {
    id: workbookId,
    name: file.name,
    appVersion: UNIVER_MODEL_VERSION,
    locale: LocaleType.KO_KR,
    styles: styleRegistry.styles,
    sheetOrder,
    sheets,
    custom: { [IMPORTED_LAYOUT_CUSTOM_KEY]: true },
  };
}

