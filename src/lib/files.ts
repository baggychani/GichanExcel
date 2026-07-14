import LuckyExcel from "@mertdeveci55/univer-import-export";
import {
  CellValueType,
  LocaleType,
  mergeWorksheetSnapshotWithDefault,
  type ICellData,
  type IWorkbookData,
} from "@univerjs/core";
import type { FUniver } from "@univerjs/core/facade";
import { invoke } from "@tauri-apps/api/core";
import { BaseDirectory } from "@tauri-apps/api/path";
import { confirm, open, save } from "@tauri-apps/plugin-dialog";
import { exists, readFile, readTextFile, remove, writeFile, writeTextFile } from "@tauri-apps/plugin-fs";
import { applySheetDefaults } from "../setup-univer";
import { UNIVER_MODEL_VERSION } from "./workbook-constants";
import { exportWorkbookToXlsx } from "./xlsx-export";
import { importExcelWithSheetJs, workbookPreservesImportedLayout } from "./xlsx-import";

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
  // 가져온 통합 문서는 원본 레이아웃을 갖고 있으므로 건드리지 않고,
  // 일반 텍스트로 만든 통합 문서에만 앱 기본 행/열 값을 적용합니다.
  if (!workbookPreservesImportedLayout(workbookData)) {
    applySheetDefaults(univerAPI);
  }
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

    importExcelWithSheetJs(file).then(resolve).catch(reject);
  });
}

function getCellText(snapshot: IWorkbookData, sheetId: string, row: number, column: number): string {
  const cell = snapshot.sheets[sheetId]?.cellData?.[row]?.[column];
  const value = cell?.v ?? "";
  return String(value);
}

function hasDelimitedExportContent(cell: ICellData | null | undefined): boolean {
  if (!cell) {
    return false;
  }

  if (typeof cell.f === "string" && cell.f.length > 0) {
    return true;
  }

  if (cell.v === undefined || cell.v === null) {
    return false;
  }

  return typeof cell.v === "string" ? cell.v.length > 0 : true;
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
  let maxRow = -1;
  let maxColumn = -1;

  Object.entries(cellData).forEach(([rowIndexText, rowCells]) => {
    const rowIndex = Number(rowIndexText);
    if (!Number.isInteger(rowIndex) || rowIndex < 0 || !rowCells) {
      return;
    }

    Object.entries(rowCells).forEach(([columnIndexText, cell]) => {
      const columnIndex = Number(columnIndexText);
      if (
        !Number.isInteger(columnIndex) ||
        columnIndex < 0 ||
        !hasDelimitedExportContent(cell as ICellData | null | undefined)
      ) {
        return;
      }

      maxRow = Math.max(maxRow, rowIndex);
      maxColumn = Math.max(maxColumn, columnIndex);
    });
  });

  if (maxRow < 0 || maxColumn < 0) {
    return "";
  }

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
    return exportDelimitedText(snapshot, ",");
  }

  return exportWorkbookToXlsx(snapshot);
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
      "CSV/TSV/TXT는 활성 시트의 값만 저장하며, 서식, 여러 시트, 수식 등의 정보가 보존되지 않을 수 있습니다. 계속 저장할까요?",
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

const RECENT_FILES_FILE = "recent-files.json";
const MAX_RECENT_FILES = 10;

export async function readRecentFiles(): Promise<string[]> {
  try {
    const hasRecent = await exists(RECENT_FILES_FILE, {
      baseDir: BaseDirectory.AppLocalData,
    });
    if (!hasRecent) {
      return [];
    }
    const content = await readTextFile(RECENT_FILES_FILE, {
      baseDir: BaseDirectory.AppLocalData,
    });
    const parsed = JSON.parse(content);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export async function addRecentFile(path: string): Promise<string[]> {
  if (!path) {
    return await readRecentFiles();
  }
  try {
    const list = await readRecentFiles();
    const filtered = list.filter((p) => p !== path);
    filtered.unshift(path);
    const updated = filtered.slice(0, MAX_RECENT_FILES);
    await writeTextFile(RECENT_FILES_FILE, JSON.stringify(updated), {
      baseDir: BaseDirectory.AppLocalData,
    });
    return updated;
  } catch {
    return [];
  }
}

export async function clearRecentFiles(): Promise<void> {
  try {
    const hasRecent = await exists(RECENT_FILES_FILE, {
      baseDir: BaseDirectory.AppLocalData,
    });
    if (hasRecent) {
      await remove(RECENT_FILES_FILE, {
        baseDir: BaseDirectory.AppLocalData,
      });
    }
  } catch {}
}
