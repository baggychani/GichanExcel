import { BaseDirectory } from "@tauri-apps/api/path";
import {
  exists,
  readTextFile,
  remove,
  rename,
  writeTextFile,
} from "@tauri-apps/plugin-fs";
import type { IWorkbookData } from "@univerjs/core";

const AUTOSAVE_FILE = "autosave.json";
const AUTOSAVE_TEMP_FILE = "autosave.json.tmp";

/**
 * Google Sheets에 가깝게: 편집이 멈춘 뒤 짧게 쉬면 복구 포인트를 남깁니다.
 * Excel AutoRecover에 가깝게: 계속 타이핑 중이어도 일정 간격마다 한 번은 남깁니다.
 */
export const AUTOSAVE_IDLE_MS = 2_000;
export const AUTOSAVE_MAX_INTERVAL_MS = 60_000;

export interface AutoSaveRecord {
  path: string | null;
  savedAt: string;
  snapshot: IWorkbookData;
}

let writeChain: Promise<void> = Promise.resolve();

export function snapshotFingerprint(snapshot: IWorkbookData): string {
  return JSON.stringify(snapshot);
}

function hasObjectValue(value: unknown): boolean {
  return !!value && typeof value === "object" && Object.keys(value).length > 0;
}

export function hasWorkbookContent(snapshot: IWorkbookData): boolean {
  return Object.values(snapshot.sheets ?? {}).some((sheet) => {
    return (
      hasObjectValue(sheet.cellData) ||
      hasObjectValue(sheet.mergeData) ||
      hasObjectValue(sheet.rowData) ||
      hasObjectValue(sheet.columnData)
    );
  });
}

export async function writeAutoSave(record: AutoSaveRecord): Promise<void> {
  const payload = JSON.stringify(record);

  writeChain = writeChain
    .catch(() => undefined)
    .then(async () => {
      await writeTextFile(AUTOSAVE_TEMP_FILE, payload, {
        baseDir: BaseDirectory.AppLocalData,
      });
      await rename(AUTOSAVE_TEMP_FILE, AUTOSAVE_FILE, {
        oldPathBaseDir: BaseDirectory.AppLocalData,
        newPathBaseDir: BaseDirectory.AppLocalData,
      });
    });

  await writeChain;
}

export async function readAutoSave(): Promise<AutoSaveRecord | null> {
  const hasAutoSave = await exists(AUTOSAVE_FILE, {
    baseDir: BaseDirectory.AppLocalData,
  });

  if (!hasAutoSave) {
    return null;
  }

  const content = await readTextFile(AUTOSAVE_FILE, {
    baseDir: BaseDirectory.AppLocalData,
  });

  return JSON.parse(content) as AutoSaveRecord;
}

export async function clearAutoSave(): Promise<void> {
  writeChain = writeChain.catch(() => undefined).then(async () => {
    const hasAutoSave = await exists(AUTOSAVE_FILE, {
      baseDir: BaseDirectory.AppLocalData,
    });
    if (hasAutoSave) {
      await remove(AUTOSAVE_FILE, {
        baseDir: BaseDirectory.AppLocalData,
      });
    }

    const hasTemp = await exists(AUTOSAVE_TEMP_FILE, {
      baseDir: BaseDirectory.AppLocalData,
    });
    if (hasTemp) {
      await remove(AUTOSAVE_TEMP_FILE, {
        baseDir: BaseDirectory.AppLocalData,
      });
    }
  });

  await writeChain;
}

export function didSnapshotChange(
  previousFingerprint: string | null,
  snapshot: IWorkbookData,
): { changed: boolean; fingerprint: string } {
  const fingerprint = snapshotFingerprint(snapshot);
  return {
    changed: previousFingerprint !== fingerprint,
    fingerprint,
  };
}

/**
 * dirty 판정용 문서 지문.
 * - 포함: 셀 값/수식/스타일, 병합, 행·열 크기 등 사용자가 바꾼 문서 상태
 * - 제외: 선택/스크롤 같은 UI 상태(스냅샷에 원래 없음)
 */
export function documentFingerprint(snapshot: IWorkbookData): string {
  const sheets = Object.entries(snapshot.sheets ?? {}).map(([sheetId, sheet]) => ({
    id: sheetId,
    name: sheet.name ?? null,
    cellData: sheet.cellData ?? {},
    mergeData: sheet.mergeData ?? [],
    rowData: sheet.rowData ?? {},
    columnData: sheet.columnData ?? {},
    rowCount: sheet.rowCount ?? null,
    columnCount: sheet.columnCount ?? null,
    defaultStyle: sheet.defaultStyle ?? null,
  }));

  return JSON.stringify({
    sheetOrder: snapshot.sheetOrder ?? [],
    styles: snapshot.styles ?? {},
    sheets,
  });
}
