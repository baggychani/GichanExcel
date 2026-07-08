import LuckyExcel from "@mertdeveci55/univer-import-export";
import type { IWorkbookData } from "@univerjs/core";
import type { FUniver } from "@univerjs/core/facade";
import { open, save } from "@tauri-apps/plugin-dialog";
import { readFile, writeFile } from "@tauri-apps/plugin-fs";

const SPREADSHEET_FILTERS = [
  {
    name: "스프레드시트",
    extensions: ["xlsx", "xls", "csv"],
  },
  {
    name: "Excel",
    extensions: ["xlsx", "xls"],
  },
  {
    name: "CSV",
    extensions: ["csv"],
  },
];

export interface DocumentState {
  path: string | null;
  dirty: boolean;
}

function getExtension(path: string): string {
  const index = path.lastIndexOf(".");
  return index >= 0 ? path.slice(index + 1).toLowerCase() : "";
}

function getFileName(path: string): string {
  const normalized = path.replace(/\\/g, "/");
  return normalized.slice(normalized.lastIndexOf("/") + 1);
}

function mimeForExtension(ext: string): string {
  switch (ext) {
    case "csv":
      return "text/csv";
    case "xls":
      return "application/vnd.ms-excel";
    default:
      return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }
}

async function pathToFile(path: string): Promise<File> {
  const bytes = await readFile(path);
  const ext = getExtension(path);
  return new File([bytes], getFileName(path), { type: mimeForExtension(ext) });
}

function loadWorkbookData(
  univerAPI: FUniver,
  workbookData: IWorkbookData,
): void {
  const activeWorkbook = univerAPI.getActiveWorkbook();
  if (activeWorkbook) {
    univerAPI.disposeUnit(activeWorkbook.getId());
  }

  univerAPI.createWorkbook(workbookData);
  univerAPI.getFormula().executeCalculation();
}

function importFile(file: File): Promise<IWorkbookData> {
  const ext = getExtension(file.name);

  return new Promise((resolve, reject) => {
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

async function exportToBuffer(
  snapshot: IWorkbookData,
  ext: "xlsx" | "csv",
): Promise<ArrayBuffer | Uint8Array | string> {
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

async function writeSnapshot(path: string, snapshot: IWorkbookData): Promise<void> {
  const ext = getExtension(path) as "xlsx" | "csv" | "xls";
  const exportExt = ext === "csv" ? "csv" : "xlsx";
  const payload = await exportToBuffer(snapshot, exportExt);

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
    filters: SPREADSHEET_FILTERS,
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
  _currentPath: string | null,
): Promise<DocumentState> {
  const workbook = univerAPI.getActiveWorkbook();
  if (!workbook) {
    throw new Error("열린 통합 문서가 없습니다.");
  }

  const selected = await save({
    filters: SPREADSHEET_FILTERS,
    defaultPath: "새 통합문서.xlsx",
  });

  if (!selected) {
    throw new Error("cancelled");
  }

  const snapshot = workbook.save();
  await writeSnapshot(selected, snapshot);

  return { path: selected, dirty: false };
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

export function getWindowTitle(path: string | null, dirty: boolean): string {
  const prefix = dirty ? "● " : "";
  if (!path) {
    return `${prefix}기찬엑셀`;
  }

  return `${prefix}${getFileName(path)} - 기찬엑셀`;
}
