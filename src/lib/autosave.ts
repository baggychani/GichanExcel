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
const AUTOSAVE_CHANNEL = "gichan-excel-autosave-owner";
const AUTOSAVE_TEMP_PREFIX = "autosave.";
const AUTOSAVE_TEMP_SUFFIX = ".json.tmp";

/**
 * Google Sheets에 가깝게: 편집이 멈춘 뒤 짧게 쉬면 복구 포인트를 남깁니다.
 * Excel AutoRecover에 가깝게: 계속 타이핑 중이어도 일정 간격마다 한 번은 남깁니다.
 */
export const AUTOSAVE_IDLE_MS = 2_000;
export const AUTOSAVE_MAX_INTERVAL_MS = 60_000;

export interface AutoSaveRecord {
  path: string | null;
  ownerId?: string;
  savedAt: string;
  snapshot: IWorkbookData;
}

let writeChain: Promise<void> = Promise.resolve();

function getTempFileName(ownerId: string | undefined): string {
  const safeOwnerId = (ownerId ?? "legacy").replace(/[^a-z0-9_-]/gi, "_");
  return `${AUTOSAVE_TEMP_PREFIX}${safeOwnerId}${AUTOSAVE_TEMP_SUFFIX}`;
}

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
  const tempFile = getTempFileName(record.ownerId);

  writeChain = writeChain
    .catch(() => undefined)
    .then(async () => {
      await writeTextFile(tempFile, payload, {
        baseDir: BaseDirectory.AppLocalData,
      });
      await rename(tempFile, AUTOSAVE_FILE, {
        oldPathBaseDir: BaseDirectory.AppLocalData,
        newPathBaseDir: BaseDirectory.AppLocalData,
      });
    });

  await writeChain;
}

export async function readAutoSave(): Promise<AutoSaveRecord | null> {
  try {
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
  } catch {
    return null;
  }
}

interface ClearAutoSaveOptions {
  force?: boolean;
  ownerId?: string;
  path?: string | null;
}

export async function clearAutoSave(options: ClearAutoSaveOptions = {}): Promise<void> {
  writeChain = writeChain.catch(() => undefined).then(async () => {
    const hasAutoSave = await exists(AUTOSAVE_FILE, {
      baseDir: BaseDirectory.AppLocalData,
    });
    if (hasAutoSave) {
      if (!options.force) {
        const record = await readAutoSave();
        if (!record) {
          await remove(AUTOSAVE_FILE, {
            baseDir: BaseDirectory.AppLocalData,
          });
          return;
        }
        const ownerMatches = !!(
          options.ownerId &&
          record.ownerId &&
          record.ownerId === options.ownerId
        );
        if (
          options.ownerId &&
          record.ownerId &&
          record.ownerId !== options.ownerId
        ) {
          return;
        }
        if (!ownerMatches && "path" in options && record.path !== options.path) {
          return;
        }
      }
      await remove(AUTOSAVE_FILE, {
        baseDir: BaseDirectory.AppLocalData,
      });
    }

    const tempFile = getTempFileName(options.ownerId);
    const hasTemp = await exists(tempFile, {
      baseDir: BaseDirectory.AppLocalData,
    });
    if (hasTemp) {
      await remove(tempFile, {
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

type AutoSaveOwnerMessage =
  | {
      type: "autosave-owner-ping";
      ownerId: string;
      requestId: string;
    }
  | {
      type: "autosave-owner-pong";
      ownerId: string;
      requestId: string;
    };

export function announceAutoSaveOwner(ownerId: string): () => void {
  if (typeof BroadcastChannel === "undefined") {
    return () => undefined;
  }

  const channel = new BroadcastChannel(AUTOSAVE_CHANNEL);
  channel.addEventListener("message", (event: MessageEvent<AutoSaveOwnerMessage>) => {
    const message = event.data;
    if (message?.type !== "autosave-owner-ping" || message.ownerId !== ownerId) {
      return;
    }

    channel.postMessage({
      type: "autosave-owner-pong",
      ownerId,
      requestId: message.requestId,
    } satisfies AutoSaveOwnerMessage);
  });

  return () => channel.close();
}

export function isAutoSaveOwnerActive(
  ownerId: string,
  timeoutMs = 500,
): Promise<boolean> {
  if (typeof BroadcastChannel === "undefined") {
    return Promise.resolve(false);
  }

  return new Promise((resolve) => {
    const channel = new BroadcastChannel(AUTOSAVE_CHANNEL);
    const requestId = crypto.randomUUID();
    const timer = window.setTimeout(() => {
      channel.close();
      resolve(false);
    }, timeoutMs);

    channel.addEventListener("message", (event: MessageEvent<AutoSaveOwnerMessage>) => {
      const message = event.data;
      if (
        message?.type !== "autosave-owner-pong" ||
        message.ownerId !== ownerId ||
        message.requestId !== requestId
      ) {
        return;
      }

      window.clearTimeout(timer);
      channel.close();
      resolve(true);
    });

    channel.postMessage({
      type: "autosave-owner-ping",
      ownerId,
      requestId,
    } satisfies AutoSaveOwnerMessage);
  });
}
