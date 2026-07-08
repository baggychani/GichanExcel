import { BaseDirectory } from "@tauri-apps/api/path";
import {
  exists,
  readTextFile,
  remove,
  writeTextFile,
} from "@tauri-apps/plugin-fs";
import type { IWorkbookData } from "@univerjs/core";

const AUTOSAVE_FILE = "autosave.json";

export interface AutoSaveRecord {
  path: string | null;
  savedAt: string;
  snapshot: IWorkbookData;
}

export async function writeAutoSave(record: AutoSaveRecord): Promise<void> {
  await writeTextFile(AUTOSAVE_FILE, JSON.stringify(record), {
    baseDir: BaseDirectory.AppLocalData,
  });
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
  const hasAutoSave = await exists(AUTOSAVE_FILE, {
    baseDir: BaseDirectory.AppLocalData,
  });

  if (!hasAutoSave) {
    return;
  }

  await remove(AUTOSAVE_FILE, {
    baseDir: BaseDirectory.AppLocalData,
  });
}
