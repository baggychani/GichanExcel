import LuckyExcel from "@mertdeveci55/univer-import-export";
import {
  CellValueType,
  LocaleType,
  mergeWorksheetSnapshotWithDefault,
  type IWorkbookData,
} from "@univerjs/core";
import type { FUniver } from "@univerjs/core/facade";
import { invoke } from "@tauri-apps/api/core";
import { confirm, open, save } from "@tauri-apps/plugin-dialog";
import { readFile, writeFile } from "@tauri-apps/plugin-fs";
import { applySheetDefaults } from "../setup-univer";

const OPEN_SPREADSHEET_FILTERS = [
  {
    name: "스프레드시트",
    extensions: ["xlsx", "xls", "csv", "tsv", "txt"],
  },
  {
    name: "Excel",
    extensions: ["xlsx", "xls"],
  },
  {
    name: "CSV",
    extensions: ["csv"],
  },
  {
    name: "탭 구분 텍스트",
    extensions: ["tsv", "txt"],
  },
];

const SAVE_SPREADSHEET_FILTERS = [
  {
    name: "Excel 통합 문서",
    extensions: ["xlsx"],
  },
  {
    name: "CSV",
    extensions: ["csv"],
  },
  {
    name: "탭 구분 텍스트",
    extensions: ["tsv", "txt"],
  },
];

const TEXT_FORMAT_EXTENSIONS = new Set(["csv", "tsv", "txt"]);
const UNIVER_MODEL_VERSION = "0.25.1";

export interface DocumentState {
  path: string | null;
  dirty: boolean;
}

function getExtension(path: string): string {
  const index = path.lastIndexOf(".");
  return index >= 0 ? path.slice(index + 1).toLowerCase() : "";
}

export function getFileName(path: string): string {
  const normalized = path.replace(/\\/g, "/");
  return normalized.slice(normalized.lastIndexOf("/") + 1);
}

function mimeForExtension(ext: string): string {
  switch (ext) {
    case "csv":
      return "text/csv";
    case "tsv":
    case "txt":
      return "text/plain";
    case "xls":
      return "application/vnd.ms-excel";
    default:
      return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }
}

async function allowDocumentPath(path: string): Promise<void> {
  // 다이얼로그로 고른 경로는 이미 스코프에 들어가지만,
  // 파일 연결/자동저장 복구처럼 다이얼로그를 거치지 않은 경로는 별도로 허용해야 합니다.
  await invoke("allow_document_path", { path });
}

async function pathToFile(path: string): Promise<File> {
  await allowDocumentPath(path);
  const bytes = await readFile(path);
  const ext = getExtension(path);
  return new File([bytes], getFileName(path), { type: mimeForExtension(ext) });
}

export function loadWorkbookData(
  univerAPI: FUniver,
  workbookData: IWorkbookData,
): void {
  const activeWorkbook = univerAPI.getActiveWorkbook();
  if (activeWorkbook) {
    univerAPI.disposeUnit(activeWorkbook.getId());
  }

  univerAPI.createWorkbook(workbookData);
  // 새로 만들어진 워크북에도 줄바꿈 기본값과 행 높이 재계산을 다시 적용합니다.
  // (createWorkbook은 매번 새 시트를 만들기 때문에 최초 실행 시 지정한 기본값이
  //  이어지지 않습니다.)
  applySheetDefaults(univerAPI);
  univerAPI.getFormula().executeCalculation();
}

function parseDelimitedText(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let value = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === "\"") {
      if (quoted && next === "\"") {
        value += "\"";
        index += 1;
      } else {
        quoted = !quoted;
      }
      continue;
    }

    if (!quoted && char === delimiter) {
      row.push(value);
      value = "";
      continue;
    }

    if (!quoted && (char === "\n" || char === "\r")) {
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
      if (char === "\r" && next === "\n") {
        index += 1;
      }
      continue;
    }

    value += char;
  }

  row.push(value);
  if (row.some((cell) => cell.length > 0) || rows.length === 0) {
    rows.push(row);
  }

  return rows;
}

function delimitedTextToWorkbookData(
  text: string,
  delimiter: string,
  name: string,
): IWorkbookData {
  const rows = parseDelimitedText(text, delimiter);
  const workbookId = crypto.randomUUID();
  const sheetId = crypto.randomUUID();
  const columnCount = Math.max(1, ...rows.map((row) => row.length));
  const cellData: IWorkbookData["sheets"][string]["cellData"] = {};

  rows.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      if (!cell) {
        return;
      }

      cellData[rowIndex] ??= {};
      cellData[rowIndex][columnIndex] = {
        v: cell,
        t: CellValueType.STRING,
      };
    });
  });

  return {
    id: workbookId,
    name,
    appVersion: UNIVER_MODEL_VERSION,
    locale: LocaleType.KO_KR,
    styles: {},
    sheetOrder: [sheetId],
    sheets: {
      [sheetId]: mergeWorksheetSnapshotWithDefault({
        id: sheetId,
        name,
        rowCount: Math.max(100, rows.length),
        columnCount: Math.max(26, columnCount),
        cellData,
      }),
    },
  };
}

function importFile(file: File): Promise<IWorkbookData> {
  const ext = getExtension(file.name);

  return new Promise((resolve, reject) => {
    if (ext === "tsv" || ext === "txt") {
      file
        .text()
        .then((text) =>
          resolve(delimitedTextToWorkbookData(text, "\t", getFileName(file.name))),
        )
        .catch(reject);
      return;
    }

    if (ext === "csv") {
      LuckyExcel.transformCsvToUniver(
        file,
        (workbookData: IWorkbookData) => resolve(workbookData),
        (error: Error) => reject(error),
      );
      return;
    }

    LuckyExcel.transformExcelToUniver(
      file,
      (workbookData: IWorkbookData) => resolve(workbookData),
      (error: Error) => reject(error),
    ).catch(reject);
  });
}

function getCellText(snapshot: IWorkbookData, sheetId: string, row: number, column: number): string {
  const cell = snapshot.sheets[sheetId]?.cellData?.[row]?.[column];
  const value = cell?.v ?? "";
  return String(value);
}

function escapeDelimitedValue(value: string, delimiter: string): string {
  if (
    value.includes("\"") ||
    value.includes("\n") ||
    value.includes("\r") ||
    value.includes(delimiter)
  ) {
    return `"${value.replace(/"/g, "\"\"")}"`;
  }

  return value;
}

function exportDelimitedText(snapshot: IWorkbookData, delimiter: string): string {
  const sheetId = snapshot.sheetOrder[0];
  const sheet = snapshot.sheets[sheetId];
  const cellData = sheet?.cellData ?? {};
  const rowIndexes = Object.keys(cellData).map(Number);
  const maxRow = rowIndexes.length ? Math.max(...rowIndexes) : 0;
  let maxColumn = 0;

  rowIndexes.forEach((rowIndex) => {
    const columns = Object.keys(cellData[rowIndex] ?? {}).map(Number);
    if (columns.length) {
      maxColumn = Math.max(maxColumn, ...columns);
    }
  });

  const rows: string[] = [];
  for (let row = 0; row <= maxRow; row += 1) {
    const values: string[] = [];
    for (let column = 0; column <= maxColumn; column += 1) {
      values.push(escapeDelimitedValue(getCellText(snapshot, sheetId, row, column), delimiter));
    }
    rows.push(values.join(delimiter));
  }

  return rows.join("\r\n");
}

async function exportToBuffer(
  snapshot: IWorkbookData,
  ext: "xlsx" | "csv" | "tsv",
): Promise<ArrayBuffer | Uint8Array | string> {
  if (ext === "tsv") {
    return exportDelimitedText(snapshot, "\t");
  }

  if (ext === "csv") {
    return new Promise((resolve, reject) => {
      LuckyExcel.transformUniverToCsv({
        snapshot,
        getBuffer: true,
        success: (csvContent?: string | Record<string, string>) => {
          if (typeof csvContent === "string") {
            resolve(csvContent);
            return;
          }

          if (!csvContent) {
            resolve("");
            return;
          }

          const firstSheet = Object.values(csvContent)[0];
          resolve(firstSheet ?? "");
        },
        error: reject,
      }).catch(reject);
    });
  }

  return new Promise((resolve, reject) => {
    LuckyExcel.transformUniverToExcel({
      snapshot,
      getBuffer: true,
      success: (buffer?: ArrayBuffer) => {
        if (!buffer) {
          reject(new Error("내보낼 데이터가 없습니다."));
          return;
        }
        resolve(buffer);
      },
      error: reject,
    }).catch(reject);
  });
}

function normalizeSavePath(path: string): string {
  return getExtension(path) ? path : `${path}.xlsx`;
}

async function writeSnapshot(path: string, snapshot: IWorkbookData): Promise<void> {
  const ext = getExtension(path);
  if (ext === "xls") {
    throw new Error("구형 .xls 저장은 지원하지 않습니다. .xlsx로 저장해 주세요.");
  }

  const exportExt = ext === "csv" ? "csv" : ext === "tsv" || ext === "txt" ? "tsv" : "xlsx";
  const payload = await exportToBuffer(snapshot, exportExt);
  await allowDocumentPath(path);

  if (typeof payload === "string") {
    await writeFile(path, new TextEncoder().encode(payload));
    return;
  }

  const bytes = payload instanceof Uint8Array ? payload : new Uint8Array(payload);
  await writeFile(path, bytes);
}

export async function openSpreadsheet(
  univerAPI: FUniver,
): Promise<DocumentState> {
  const selected = await open({
    multiple: false,
    filters: OPEN_SPREADSHEET_FILTERS,
  });

  if (!selected || Array.isArray(selected)) {
    throw new Error("cancelled");
  }

  const file = await pathToFile(selected);
  const workbookData = await importFile(file);
  loadWorkbookData(univerAPI, workbookData);

  return { path: selected, dirty: false };
}

export async function saveSpreadsheet(
  univerAPI: FUniver,
  currentPath: string | null,
): Promise<DocumentState> {
  const workbook = univerAPI.getActiveWorkbook();
  if (!workbook) {
    throw new Error("열린 통합 문서가 없습니다.");
  }

  const snapshot = workbook.save();

  if (!currentPath) {
    return saveSpreadsheetAs(univerAPI, currentPath);
  }

  await writeSnapshot(currentPath, snapshot);
  return { path: currentPath, dirty: false };
}

export async function saveSpreadsheetAs(
  univerAPI: FUniver,
  currentPath: string | null,
): Promise<DocumentState> {
  const workbook = univerAPI.getActiveWorkbook();
  if (!workbook) {
    throw new Error("열린 통합 문서가 없습니다.");
  }

  // Excel과 같이: 이미 저장된 파일이면 현재 이름을 기본값으로 둡니다.
  const selected = await save({
    filters: SAVE_SPREADSHEET_FILTERS,
    defaultPath: currentPath ? getFileName(currentPath) : "새 통합문서.xlsx",
  });

  if (!selected) {
    throw new Error("cancelled");
  }

  const normalizedPath = normalizeSavePath(selected);
  const ext = getExtension(normalizedPath);
  if (TEXT_FORMAT_EXTENSIONS.has(ext)) {
    const accepted = await confirm(
      "CSV/TSV/TXT는 활성 시트의 값만 저장하며 서식, 여러 시트, 수식 일부 정보가 보존되지 않을 수 있습니다. 계속 저장할까요?",
      {
        title: "텍스트 형식으로 저장",
        kind: "warning",
        okLabel: "저장",
        cancelLabel: "취소",
      },
    );

    if (!accepted) {
      throw new Error("cancelled");
    }
  }

  const snapshot = workbook.save();
  await writeSnapshot(normalizedPath, snapshot);

  return { path: normalizedPath, dirty: false };
}

export async function openPath(
  univerAPI: FUniver,
  path: string,
): Promise<DocumentState> {
  const file = await pathToFile(path);
  const workbookData = await importFile(file);
  loadWorkbookData(univerAPI, workbookData);
  return { path, dirty: false };
}

export async function openFileBytes(
  univerAPI: FUniver,
  path: string,
  bytes: number[] | Uint8Array,
): Promise<DocumentState> {
  const ext = getExtension(path);
  const payload = bytes instanceof Uint8Array ? bytes : Uint8Array.from(bytes);
  const file = new File([payload], getFileName(path), {
    type: mimeForExtension(ext),
  });
  const workbookData = await importFile(file);
  loadWorkbookData(univerAPI, workbookData);
  return { path, dirty: false };
}

export function getWindowTitle(path: string | null, dirty: boolean): string {
  const prefix = dirty ? "● " : "";
  if (!path) {
    return `${prefix}기찬엑셀`;
  }

  return `${prefix}${getFileName(path)} - 기찬엑셀`;
}
