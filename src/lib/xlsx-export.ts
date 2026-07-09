import ExcelJS from "exceljs";
import {
  BooleanNumber,
  BorderStyleTypes,
  CellValueType,
  HorizontalAlign,
  VerticalAlign,
  WrapStrategy,
  type IBorderStyleData,
  type ICellData,
  type IStyleData,
  type IWorkbookData,
} from "@univerjs/core";

function isTruthyBooleanNumber(value: unknown): boolean {
  return value === BooleanNumber.TRUE || value === true || value === 1;
}

function toArgb(color: string | null | undefined): string | undefined {
  if (!color) {
    return undefined;
  }

  const hex = color.replace(/[^0-9a-f]/gi, "").toUpperCase();
  if (hex.length === 6) {
    return `FF${hex}`;
  }
  if (hex.length === 8) {
    return hex;
  }
  return undefined;
}

function resolveStyle(
  styleRef: ICellData["s"] | undefined,
  workbookStyles: IWorkbookData["styles"] | undefined,
): IStyleData | undefined {
  if (!styleRef) {
    return undefined;
  }
  if (typeof styleRef === "string") {
    const resolved = workbookStyles?.[styleRef];
    return resolved && typeof resolved === "object" ? resolved : undefined;
  }
  return styleRef;
}

function mapBorderStyle(style: BorderStyleTypes | undefined): ExcelJS.BorderStyle | undefined {
  switch (style) {
    case BorderStyleTypes.THIN:
      return "thin";
    case BorderStyleTypes.HAIR:
      return "hair";
    case BorderStyleTypes.DOTTED:
      return "dotted";
    case BorderStyleTypes.DASHED:
      return "dashed";
    case BorderStyleTypes.DASH_DOT:
      return "dashDot";
    case BorderStyleTypes.DASH_DOT_DOT:
      return "dashDotDot";
    case BorderStyleTypes.DOUBLE:
      return "double";
    case BorderStyleTypes.MEDIUM:
      return "medium";
    case BorderStyleTypes.MEDIUM_DASHED:
      return "mediumDashed";
    case BorderStyleTypes.MEDIUM_DASH_DOT:
      return "mediumDashDot";
    case BorderStyleTypes.MEDIUM_DASH_DOT_DOT:
      return "mediumDashDotDot";
    case BorderStyleTypes.SLANT_DASH_DOT:
      return "slantDashDot";
    case BorderStyleTypes.THICK:
      return "thick";
    default:
      return undefined;
  }
}

function mapBorderSide(side: IBorderStyleData | null | undefined): Partial<ExcelJS.Border> | undefined {
  if (!side) {
    return undefined;
  }
  const style = mapBorderStyle(side.s);
  if (!style) {
    return undefined;
  }
  const argb = toArgb(side.cl?.rgb ?? undefined);
  return {
    style,
    ...(argb ? { color: { argb } } : {}),
  };
}

function applyStyleToCell(cell: ExcelJS.Cell, style: IStyleData | undefined): void {
  if (!style) {
    return;
  }

  const font: Partial<ExcelJS.Font> = {};
  if (style.ff) {
    font.name = style.ff;
  }
  if (typeof style.fs === "number") {
    font.size = style.fs;
  }
  if (isTruthyBooleanNumber(style.bl)) {
    font.bold = true;
  }
  if (isTruthyBooleanNumber(style.it)) {
    font.italic = true;
  }
  if (style.ul?.s === BooleanNumber.TRUE || isTruthyBooleanNumber(style.ul?.s)) {
    font.underline = true;
  }
  if (style.st?.s === BooleanNumber.TRUE || isTruthyBooleanNumber(style.st?.s)) {
    font.strike = true;
  }
  const fontColor = toArgb(style.cl?.rgb ?? undefined);
  if (fontColor) {
    font.color = { argb: fontColor };
  }
  if (Object.keys(font).length) {
    cell.font = font;
  }

  const fillArgb = toArgb(style.bg?.rgb ?? undefined);
  if (fillArgb) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: fillArgb },
    };
  }

  if (style.bd) {
    const border: Partial<ExcelJS.Borders> = {};
    const top = mapBorderSide(style.bd.t ?? undefined);
    const left = mapBorderSide(style.bd.l ?? undefined);
    const bottom = mapBorderSide(style.bd.b ?? undefined);
    const right = mapBorderSide(style.bd.r ?? undefined);
    if (top) {
      border.top = top;
    }
    if (left) {
      border.left = left;
    }
    if (bottom) {
      border.bottom = bottom;
    }
    if (right) {
      border.right = right;
    }
    if (Object.keys(border).length) {
      cell.border = border;
    }
  }

  const alignment: Partial<ExcelJS.Alignment> = {};
  switch (style.ht) {
    case HorizontalAlign.LEFT:
      alignment.horizontal = "left";
      break;
    case HorizontalAlign.CENTER:
      alignment.horizontal = "center";
      break;
    case HorizontalAlign.RIGHT:
      alignment.horizontal = "right";
      break;
    case HorizontalAlign.JUSTIFIED:
      alignment.horizontal = "justify";
      break;
    case HorizontalAlign.DISTRIBUTED:
      alignment.horizontal = "distributed";
      break;
    default:
      break;
  }
  switch (style.vt) {
    case VerticalAlign.TOP:
      alignment.vertical = "top";
      break;
    case VerticalAlign.MIDDLE:
      alignment.vertical = "middle";
      break;
    case VerticalAlign.BOTTOM:
      alignment.vertical = "bottom";
      break;
    default:
      break;
  }
  if (style.tb === WrapStrategy.WRAP) {
    alignment.wrapText = true;
  }
  if (Object.keys(alignment).length) {
    cell.alignment = alignment;
  }

  if (style.n?.pattern) {
    cell.numFmt = style.n.pattern;
  }
}

function applyCellValue(target: ExcelJS.Cell, cell: ICellData): void {
  if (cell.f) {
    const formula = cell.f.startsWith("=") ? cell.f.slice(1) : cell.f;
    target.value = {
      formula,
      ...(cell.v !== undefined && cell.v !== null ? { result: cell.v as string | number | boolean } : {}),
    };
    return;
  }

  if (cell.v === undefined || cell.v === null) {
    return;
  }

  if (cell.t === CellValueType.NUMBER || typeof cell.v === "number") {
    target.value = Number(cell.v);
    return;
  }

  if (cell.t === CellValueType.BOOLEAN || typeof cell.v === "boolean") {
    target.value = Boolean(cell.v);
    return;
  }

  target.value = String(cell.v);
}

function columnLetter(index: number): string {
  let n = index + 1;
  let result = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

export async function exportWorkbookToXlsx(snapshot: IWorkbookData): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "GichanExcel";
  workbook.created = new Date();
  workbook.modified = new Date();

  const sheetIds = snapshot.sheetOrder?.length
    ? snapshot.sheetOrder
    : Object.keys(snapshot.sheets ?? {});

  if (sheetIds.length === 0) {
    workbook.addWorksheet("Sheet1");
  }

  for (const sheetId of sheetIds) {
    const sheet = snapshot.sheets?.[sheetId];
    if (!sheet) {
      continue;
    }

    const worksheet = workbook.addWorksheet(sheet.name || "Sheet1");
    const cellData = sheet.cellData ?? {};
    const rowData = sheet.rowData ?? {};
    const columnData = sheet.columnData ?? {};

    Object.entries(columnData).forEach(([columnIndexText, columnInfo]) => {
      const columnIndex = Number(columnIndexText);
      if (!Number.isInteger(columnIndex) || columnIndex < 0 || !columnInfo) {
        return;
      }
      const column = worksheet.getColumn(columnIndex + 1);
      if (typeof columnInfo.w === "number" && columnInfo.w > 0) {
        // Univer 픽셀 너비 ≈ Excel 문자 너비 * ~7~8. ExcelJS width는 문자 단위.
        column.width = Math.max(1, columnInfo.w / 7);
      }
      if (columnInfo.hd === BooleanNumber.TRUE) {
        column.hidden = true;
      }
    });

    Object.entries(rowData).forEach(([rowIndexText, rowInfo]) => {
      const rowIndex = Number(rowIndexText);
      if (!Number.isInteger(rowIndex) || rowIndex < 0 || !rowInfo) {
        return;
      }
      const row = worksheet.getRow(rowIndex + 1);
      if (typeof rowInfo.h === "number" && rowInfo.h > 0) {
        row.height = rowInfo.h;
      }
      if (rowInfo.hd === BooleanNumber.TRUE) {
        row.hidden = true;
      }
    });

    Object.entries(cellData).forEach(([rowIndexText, rowCells]) => {
      const rowIndex = Number(rowIndexText);
      if (!Number.isInteger(rowIndex) || rowIndex < 0 || !rowCells) {
        return;
      }

      Object.entries(rowCells).forEach(([columnIndexText, rawCell]) => {
        const columnIndex = Number(columnIndexText);
        const cell = rawCell as ICellData | null | undefined;
        if (!Number.isInteger(columnIndex) || columnIndex < 0 || !cell) {
          return;
        }

        const target = worksheet.getCell(rowIndex + 1, columnIndex + 1);
        applyCellValue(target, cell);
        applyStyleToCell(target, resolveStyle(cell.s, snapshot.styles));
      });
    });

    (sheet.mergeData ?? []).forEach((range) => {
      const start = `${columnLetter(range.startColumn)}${range.startRow + 1}`;
      const end = `${columnLetter(range.endColumn)}${range.endRow + 1}`;
      if (start !== end) {
        worksheet.mergeCells(`${start}:${end}`);
      }
    });
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
}
