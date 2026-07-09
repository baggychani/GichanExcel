import { getCurrentWindow } from "@tauri-apps/api/window";
import { listen } from "@tauri-apps/api/event";
import { invoke } from "@tauri-apps/api/core";
import { confirm } from "@tauri-apps/plugin-dialog";
import { openUrl } from "@tauri-apps/plugin-opener";
import { useCallback, useEffect, useRef, useState } from "react";
import type { FUniver } from "@univerjs/core/facade";
import {
  getFileName,
  getWindowTitle,
  loadWorkbookData,
  openFileBytes,
  openPath,
  openSpreadsheet,
  saveSpreadsheet,
  saveSpreadsheetAs,
  type DocumentState,
} from "../lib/files";
import {
  splitActiveRange,
  type DelimiterMode,
} from "../lib/text-to-columns";
import {
  AUTOSAVE_IDLE_MS,
  AUTOSAVE_MAX_INTERVAL_MS,
  clearAutoSave,
  didSnapshotChange,
  documentFingerprint,
  hasWorkbookContent,
  readAutoSave,
  writeAutoSave,
} from "../lib/autosave";
import { checkForUpdate, type UpdateInfo } from "../lib/update-checker";
import { TEXT_SPLIT_EVENT } from "../plugins/text-split";
import { setupUniver } from "../setup-univer";
import { APP_VERSION } from "../lib/version";
import { UpdateDialog } from "./UpdateDialog";
import { UnsavedChangesDialog } from "./UnsavedChangesDialog";
import { AppLogoIcon } from "./icons";

type UnsavedChoice = "save" | "discard" | "cancel";

const INITIAL_DOC: DocumentState = { path: null, dirty: false };

interface InitialOpenFile {
  path: string;
  bytes: number[];
}

const UNDO_COMMAND_ID = "univer.command.undo";
const REDO_COMMAND_ID = "univer.command.redo";

/**
 * 시스템이 알아서 쏘는 mutation만 제외합니다.
 * - 자동 행높이: 줄바꿈 기본값 때문에 열기/클릭만으로도 발생
 * - 기본 스타일 적용: 워크북 생성 직후 내부 적용
 *
 * 사용자가 직접 한 작업은 dirty로 칩니다.
 * - 셀 값/수식 입력
 * - 굵게·색·정렬 등 스타일
 * - 행/열 수동 크기 조절
 * - 병합, 삽입/삭제 등
 */
const SYSTEM_MUTATION_IDS = new Set([
  "sheet.mutation.set-worksheet-row-auto-height",
  "sheet.mutation.set-worksheet-row-is-auto-height",
  "sheet.mutation.set-worksheet-default-style",
]);

export function SpreadsheetApp() {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<FUniver | null>(null);
  const docRef = useRef<DocumentState>(INITIAL_DOC);
  const closeAllowedRef = useRef(false);
  const documentLoadInProgressRef = useRef(false);
  const suppressDirtyRef = useRef(false);
  const unsavedChoiceResolverRef = useRef<((choice: UnsavedChoice) => void) | null>(
    null,
  );
  const lastAutoSaveFingerprintRef = useRef<string | null>(null);
  const lastAutoSaveAtRef = useRef(0);
  const autoSaveInFlightRef = useRef(false);
  const recoveryPromptActiveRef = useRef(false);
  const cleanDocumentFingerprintRef = useRef<string | null>(null);
  const [doc, setDoc] = useState<DocumentState>(INITIAL_DOC);
  const [ready, setReady] = useState(false);
  const [startupOpenChecked, setStartupOpenChecked] = useState(false);
  const [autoSaveVersion, setAutoSaveVersion] = useState(0);
  const [status, setStatus] = useState("새 통합 문서");
  const [error, setError] = useState<string | null>(null);
  const [delimiterMode, setDelimiterMode] = useState<DelimiterMode>("spaces");
  const [customDelimiter, setCustomDelimiter] = useState("");
  const [mergeDelimiters, setMergeDelimiters] = useState(true);
  const [splitDialogOpen, setSplitDialogOpen] = useState(false);
  const [updateInfo, setUpdateInfo] = useState<UpdateInfo | null>(null);
  const [unsavedDialogOpen, setUnsavedDialogOpen] = useState(false);
  const [unsavedDialogSaving, setUnsavedDialogSaving] = useState(false);

  const syncTitle = useCallback(async (state: DocumentState) => {
    const title = getWindowTitle(state.path, state.dirty);
    document.title = title;
    await getCurrentWindow().setTitle(title);
  }, []);

  const captureCleanFingerprint = useCallback(() => {
    const workbook = univerRef.current?.getActiveWorkbook();
    if (!workbook) {
      cleanDocumentFingerprintRef.current = null;
      return;
    }
    cleanDocumentFingerprintRef.current = documentFingerprint(workbook.save());
  }, []);

  const markDirty = useCallback(() => {
    if (documentLoadInProgressRef.current || suppressDirtyRef.current) {
      return;
    }

    const workbook = univerRef.current?.getActiveWorkbook();
    if (!workbook) {
      return;
    }

    const fingerprint = documentFingerprint(workbook.save());
    // 기준 스냅샷이 아직 없으면 지금 찍고 dirty로 치지 않습니다.
    // (초기 자동높이/기본스타일 적용이 끝난 뒤 기준이 잡히기 전 구간)
    if (cleanDocumentFingerprintRef.current === null) {
      cleanDocumentFingerprintRef.current = fingerprint;
      return;
    }
    // 문서 상태가 기준과 같으면(화살표 이동 등) dirty 아님.
    // 스타일·행/열 크기·값 변경은 지문이 달라지므로 dirty.
    if (fingerprint === cleanDocumentFingerprintRef.current) {
      return;
    }

    setDoc((current) => {
      if (current.dirty) {
        return current;
      }
      const next = { ...current, dirty: true };
      void syncTitle(next);
      return next;
    });
    // dirty가 이미 true여도 편집이 이어지면 idle 타이머를 다시 돌립니다.
    setAutoSaveVersion((version) => version + 1);
  }, [syncTitle]);

  const runDocumentLoad = useCallback(async <T,>(operation: () => Promise<T>): Promise<T> => {
    documentLoadInProgressRef.current = true;
    try {
      return await operation();
    } finally {
      window.setTimeout(() => {
        documentLoadInProgressRef.current = false;
      }, 250);
    }
  }, []);

  const applyDocument = useCallback(
    async (state: DocumentState, message: string) => {
      setDoc(state);
      setStatus(message);
      setError(null);
      await syncTitle(state);
      if (!state.dirty) {
        lastAutoSaveFingerprintRef.current = null;
        lastAutoSaveAtRef.current = 0;
        // 로드/저장 직후 스냅샷을 깨끗한 기준으로 잡습니다.
        window.setTimeout(() => {
          captureCleanFingerprint();
        }, 300);
        void clearAutoSave();
      }
    },
    [captureCleanFingerprint, syncTitle],
  );

  useEffect(() => {
    docRef.current = doc;
  }, [doc]);

  const saveCurrentDocument = useCallback(
    async (message: string) => {
      if (!univerRef.current) {
        return null;
      }

      const state = await saveSpreadsheet(univerRef.current, docRef.current.path);
      await applyDocument(state, message);
      return state;
    },
    [applyDocument],
  );

  const confirmUnsavedAction = useCallback(async (): Promise<boolean> => {
    if (!docRef.current.dirty) {
      return true;
    }

    setUnsavedDialogOpen(true);
    const choice = await new Promise<UnsavedChoice>((resolve) => {
      unsavedChoiceResolverRef.current = resolve;
    });

    unsavedChoiceResolverRef.current = null;

    if (choice === "cancel") {
      setUnsavedDialogOpen(false);
      return false;
    }

    if (choice === "save") {
      setUnsavedDialogSaving(true);
      try {
        await saveCurrentDocument("저장했습니다.");
      } catch (err) {
        setUnsavedDialogSaving(false);
        setUnsavedDialogOpen(false);
        if (err instanceof Error && err.message === "cancelled") {
          return false;
        }
        setError(err instanceof Error ? err.message : "저장하지 못했습니다.");
        return false;
      }
      setUnsavedDialogSaving(false);
    }

    setUnsavedDialogOpen(false);
    return true;
  }, [saveCurrentDocument]);

  const handleUnsavedChoice = useCallback((choice: UnsavedChoice) => {
    unsavedChoiceResolverRef.current?.(choice);
  }, []);

  const handleOpen = useCallback(async () => {
    if (!univerRef.current) {
      return;
    }

    try {
      const canContinue = await confirmUnsavedAction();
      if (!canContinue) {
        return;
      }

      const state = await runDocumentLoad(() => openSpreadsheet(univerRef.current!));
      await applyDocument(
        state,
        state.path ? `열림: ${state.path}` : "파일을 열었습니다.",
      );
    } catch (err) {
      if (err instanceof Error && err.message === "cancelled") {
        return;
      }
      setError(err instanceof Error ? err.message : "파일을 열지 못했습니다.");
    }
  }, [applyDocument, confirmUnsavedAction, runDocumentLoad]);

  const handleSave = useCallback(async () => {
    try {
      await saveCurrentDocument("저장했습니다.");
    } catch (err) {
      if (err instanceof Error && err.message === "cancelled") {
        return;
      }
      setError(err instanceof Error ? err.message : "저장하지 못했습니다.");
    }
  }, [saveCurrentDocument]);

  const handleSaveAs = useCallback(async () => {
    if (!univerRef.current) {
      return;
    }

    try {
      const state = await saveSpreadsheetAs(univerRef.current, doc.path);
      await applyDocument(state, "다른 이름으로 저장했습니다.");
    } catch (err) {
      if (err instanceof Error && err.message === "cancelled") {
        return;
      }
      setError(
        err instanceof Error ? err.message : "다른 이름으로 저장하지 못했습니다.",
      );
    }
  }, [applyDocument, doc.path]);

  const openSplitDialog = useCallback(() => {
    const workbook = univerRef.current?.getActiveWorkbook();
    const activeRange = workbook?.getActiveRange();
    if (!activeRange) {
      setError("먼저 분할할 셀이나 범위를 선택해 주세요.");
      return;
    }

    setSplitDialogOpen(true);
    setError(null);
  }, []);

  const focusActiveWorkbook = useCallback(() => {
    const univerAPI = univerRef.current;
    const workbook = univerAPI?.getActiveWorkbook();
    const unitId = workbook?.getId();
    if (!univerAPI || !unitId) {
      return false;
    }

    const instanceService = (
      univerAPI as unknown as {
        _univerInstanceService?: {
          focusUnit: (unitId: string | null) => void;
        };
      }
    )._univerInstanceService;
    instanceService?.focusUnit(unitId);
    return true;
  }, []);

  const runWithoutDirty = useCallback(<T,>(operation: () => T): T => {
    suppressDirtyRef.current = true;
    try {
      return operation();
    } finally {
      window.setTimeout(() => {
        suppressDirtyRef.current = false;
      }, 0);
    }
  }, []);

  const stabilizeSpreadsheetFocus = useCallback(() => {
    const attempts = [0, 50, 150, 350, 800];
    const timers = attempts.map((delay) =>
      window.setTimeout(() => {
        void getCurrentWindow().setFocus().catch(() => undefined);

        const container = containerRef.current;
        const univerAPI = univerRef.current;
        const workbook = univerAPI?.getActiveWorkbook();
        const sheet = workbook?.getActiveSheet();

        runWithoutDirty(() => {
          if (workbook && sheet && !workbook.getActiveRange()) {
            workbook.setActiveRange(sheet.getRange(0, 0));
          }
        });

        focusActiveWorkbook();

        const target =
          container?.querySelector<HTMLElement>(
            '[data-u-comp="editor"], [data-u-comp="render-canvas"], [data-u-comp="workbench-layout"], canvas',
          ) ?? container;
        const active = document.activeElement;
        if (
          target &&
          (!active || active === document.body || active === document.documentElement)
        ) {
          if (!target.hasAttribute("tabindex")) {
            target.setAttribute("tabindex", "-1");
          }
          target.focus({ preventScroll: true });
        }
      }, delay),
    );

    return () => timers.forEach((timer) => window.clearTimeout(timer));
  }, [focusActiveWorkbook, runWithoutDirty]);

  const handleSplitSelection = useCallback(() => {
    if (!univerRef.current) {
      return;
    }

    try {
      const range = splitActiveRange(univerRef.current, {
        mode: delimiterMode,
        customDelimiter,
        treatMultipleDelimitersAsOne: mergeDelimiters,
      });
      setStatus(`${range} 범위를 열로 분할했습니다.`);
      setError(null);
      setSplitDialogOpen(false);
      markDirty();
    } catch (err) {
      setError(err instanceof Error ? err.message : "텍스트를 분할하지 못했습니다.");
    }
  }, [customDelimiter, delimiterMode, markDirty, mergeDelimiters]);

  const runUndoRedo = useCallback(async (mode: "undo" | "redo") => {
    const univerAPI = univerRef.current;
    if (!univerAPI) {
      return;
    }

    focusActiveWorkbook();
    const commandId = mode === "redo" ? REDO_COMMAND_ID : UNDO_COMMAND_ID;
    const didRun = univerAPI.syncExecuteCommand<object, boolean>(commandId);
    if (!didRun) {
      await (mode === "redo" ? univerAPI.redo() : univerAPI.undo());
    }
  }, [focusActiveWorkbook]);

  useEffect(() => {
    // Univer 내장 Ctrl+Z/Ctrl+Y 단축키는 document.activeElement가 특정
    // data-u-comp="editor" 속성을 가진 요소일 때만 동작하도록 되어 있는데,
    // Tauri 웹뷰에서는 이 포커스 감지가 불안정해서 되돌리기/다시 실행이
    // 아예 먹통이 되는 경우가 있습니다. 활성 워크북의 undo()/redo()는 이런
    // 포커스 조건 없이 바로 스프레드시트 편집 스택을 실행하므로, 여기서 직접
    // 처리하고 내장 단축키 처리기가 같은 키 입력을 다시 실행하지 않게 막습니다.
    // (이 effect가 setupUniver보다 먼저 등록되어야 캡처 단계에서 먼저 실행됩니다.)
    const isEditableFocus = () => {
      const active = document.activeElement as HTMLElement | null;
      if (!active) {
        return false;
      }

      const tag = active.tagName;
      return (
        tag === "INPUT" ||
        tag === "TEXTAREA" ||
        active.isContentEditable ||
        !!active.closest('input, textarea, [contenteditable="true"]')
      );
    };

    const onUndoRedoKeyDown = (event: KeyboardEvent) => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey) {
        return;
      }
      if (event.isComposing || event.keyCode === 229) {
        return;
      }

      const key = event.key.toLowerCase();
      const isUndo = key === "z" && !event.shiftKey;
      const isRedo = (key === "z" && event.shiftKey) || key === "y";
      if (!isUndo && !isRedo) {
        return;
      }

      // 우리 앱 다이얼로그의 입력창(예: 구분자 직접 입력)에 포커스가 있으면
      // 브라우저 기본 되돌리기 동작을 그대로 둡니다.
      if (isEditableFocus()) {
        return;
      }

      if (!univerRef.current) {
        return;
      }

      event.preventDefault();
      event.stopImmediatePropagation();
      void runUndoRedo(isRedo ? "redo" : "undo");
    };

    window.addEventListener("keydown", onUndoRedoKeyDown, true);
    return () => window.removeEventListener("keydown", onUndoRedoKeyDown, true);
  }, [runUndoRedo]);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) {
      return;
    }

    const onFocusWorkbook = () => {
      focusActiveWorkbook();
    };

    container.addEventListener("focusin", onFocusWorkbook, true);
    container.addEventListener("pointerdown", onFocusWorkbook, true);
    return () => {
      container.removeEventListener("focusin", onFocusWorkbook, true);
      container.removeEventListener("pointerdown", onFocusWorkbook, true);
    };
  }, [focusActiveWorkbook]);

  useEffect(() => {
    if (!containerRef.current) {
      return;
    }

    const app = setupUniver("univer-container");
    univerRef.current = app.univerAPI;
    const cancelFocusStabilization = stabilizeSpreadsheetFocus();
    setReady(true);

    return () => {
      cancelFocusStabilization();
      setReady(false);
      app.dispose();
      univerRef.current = null;
    };
  }, [stabilizeSpreadsheetFocus]);

  useEffect(() => {
    if (!ready || !univerRef.current) {
      return;
    }

    let cancelled = false;
    void invoke<InitialOpenFile | null>("initial_open_file")
      .then(async (file) => {
        if (cancelled || !univerRef.current) {
          return;
        }

        if (!file) {
          return;
        }

        const state = await runDocumentLoad(() =>
          openFileBytes(univerRef.current!, file.path, file.bytes),
        );
        if (cancelled) {
          return;
        }
        stabilizeSpreadsheetFocus();
        await applyDocument(state, `열림: ${state.path ?? "알 수 없는 파일"}`);
      })
      .catch((err) => {
        if (!cancelled) {
          setError(
            err instanceof Error ? err.message : "시작 파일을 열지 못했습니다.",
          );
        }
      })
      .finally(() => {
        if (!cancelled) {
          setStartupOpenChecked(true);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [applyDocument, ready, runDocumentLoad, stabilizeSpreadsheetFocus]);

  useEffect(() => {
    if (!ready || !startupOpenChecked || !univerRef.current) {
      return;
    }

    let cancelled = false;
    recoveryPromptActiveRef.current = true;
    void readAutoSave()
      .then(async (record) => {
        if (!record || cancelled || !univerRef.current) {
          return;
        }

        // 파일 연결로 이미 문서를 연 뒤에는, 같은 파일의 복구본만 제안합니다.
        if (docRef.current.path && record.path !== docRef.current.path) {
          return;
        }

        const savedAt = new Date(record.savedAt).toLocaleString();
        const restore = await confirm(
          `자동저장된 복구본이 있습니다.\n저장 시각: ${savedAt}\n복구할까요?`,
          {
            title: "자동저장 복구",
            kind: "warning",
            okLabel: "복구",
            cancelLabel: "삭제",
          },
        );

        if (cancelled || !univerRef.current) {
          return;
        }

        if (!restore) {
          await clearAutoSave();
          return;
        }

        await runDocumentLoad(async () => {
          loadWorkbookData(univerRef.current!, record.snapshot);
        });
        stabilizeSpreadsheetFocus();
        const { fingerprint } = didSnapshotChange(null, record.snapshot);
        lastAutoSaveFingerprintRef.current = fingerprint;
        lastAutoSaveAtRef.current = Date.now();
        await applyDocument(
          { path: record.path, dirty: true },
          `자동저장 복구본을 열었습니다. (${savedAt})`,
        );
      })
      .catch(() => {
        setError("자동저장 복구본을 읽지 못했습니다.");
      })
      .finally(() => {
        recoveryPromptActiveRef.current = false;
      });

    return () => {
      cancelled = true;
      recoveryPromptActiveRef.current = false;
    };
  }, [applyDocument, ready, runDocumentLoad, stabilizeSpreadsheetFocus, startupOpenChecked]);

  useEffect(() => {
    const onOpenTextSplit = () => openSplitDialog();
    window.addEventListener(TEXT_SPLIT_EVENT, onOpenTextSplit);
    return () => window.removeEventListener(TEXT_SPLIT_EVENT, onOpenTextSplit);
  }, [openSplitDialog]);

  useEffect(() => {
    if (!ready || !univerRef.current) {
      return;
    }

    // 초기 워크북 생성·기본 스타일 적용이 끝난 뒤 깨끗한 기준을 찍습니다.
    const timer = window.setTimeout(() => {
      if (!docRef.current.dirty) {
        captureCleanFingerprint();
      }
    }, 400);

    const disposable = univerRef.current.addEvent(
      univerRef.current.Event.CommandExecuted,
      (event) => {
        // 선택/화살표 이동은 OPERATION이라 여기 안 옴.
        if (event.type !== 2) {
          return;
        }

        // 자동높이·기본스타일처럼 시스템이 쏘는 mutation은 dirty로 치지 않고,
        // 아직 깨끗한 문서면 기준 지문만 최신으로 갱신합니다.
        if (SYSTEM_MUTATION_IDS.has(event.id)) {
          if (!docRef.current.dirty) {
            captureCleanFingerprint();
          }
          return;
        }

        markDirty();
      },
    );

    return () => {
      window.clearTimeout(timer);
      disposable.dispose();
    };
  }, [captureCleanFingerprint, markDirty, ready]);

  useEffect(() => {
    const window = getCurrentWindow();
    const unlistenPromise = window.onCloseRequested(async (event) => {
      if (closeAllowedRef.current || !docRef.current.dirty) {
        return;
      }

      event.preventDefault();
      if (unsavedChoiceResolverRef.current || recoveryPromptActiveRef.current) {
        return;
      }

      const canClose = await confirmUnsavedAction();

      if (!canClose) {
        return;
      }

      closeAllowedRef.current = true;
      try {
        await clearAutoSave();
        await window.destroy();
      } catch {
        // destroy()가 실패해도 창이 멈춰있지 않도록 강제로 닫기를 다시 시도합니다.
        closeAllowedRef.current = false;
        setError("창을 닫지 못했습니다. 다시 시도해 주세요.");
      }
    });

    return () => {
      void unlistenPromise.then((unlisten) => unlisten());
    };
  }, [confirmUnsavedAction]);

  useEffect(() => {
    if (!doc.dirty || !univerRef.current) {
      return;
    }

    const persistRecoveryPoint = () => {
      if (autoSaveInFlightRef.current || recoveryPromptActiveRef.current) {
        return;
      }

      const workbook = univerRef.current?.getActiveWorkbook();
      if (!workbook || !docRef.current.dirty) {
        return;
      }

      const snapshot = workbook.save();
      if (!hasWorkbookContent(snapshot)) {
        lastAutoSaveFingerprintRef.current = null;
        lastAutoSaveAtRef.current = Date.now();
        return;
      }

      const { changed, fingerprint } = didSnapshotChange(
        lastAutoSaveFingerprintRef.current,
        snapshot,
      );
      if (!changed) {
        lastAutoSaveAtRef.current = Date.now();
        return;
      }

      autoSaveInFlightRef.current = true;
      void writeAutoSave({
        path: docRef.current.path,
        savedAt: new Date().toISOString(),
        snapshot,
      })
        .then(() => {
          lastAutoSaveFingerprintRef.current = fingerprint;
          lastAutoSaveAtRef.current = Date.now();
          setStatus(`자동저장됨 ${new Date().toLocaleTimeString()}`);
        })
        .catch(() => {
          setError("자동저장을 하지 못했습니다.");
        })
        .finally(() => {
          autoSaveInFlightRef.current = false;
        });
    };

    // Google Sheets식: 입력이 멈춘 뒤 짧게 쉬면 복구 포인트.
    const idleTimer = window.setTimeout(persistRecoveryPoint, AUTOSAVE_IDLE_MS);

    // Excel AutoRecover식: 계속 편집 중이어도 최대 간격마다 한 번은 남김.
    const lastSavedAt = lastAutoSaveAtRef.current;
    const maxWait =
      lastSavedAt === 0
        ? AUTOSAVE_MAX_INTERVAL_MS
        : Math.max(0, AUTOSAVE_MAX_INTERVAL_MS - (Date.now() - lastSavedAt));
    const maxTimer = window.setTimeout(persistRecoveryPoint, maxWait);

    return () => {
      window.clearTimeout(idleTimer);
      window.clearTimeout(maxTimer);
    };
  }, [autoSaveVersion, doc.dirty]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!event.ctrlKey && !event.metaKey) {
        return;
      }

      const key = event.key.toLowerCase();
      if (key === "o") {
        event.preventDefault();
        void handleOpen();
      }
      if (key === "s" && event.shiftKey) {
        event.preventDefault();
        void handleSaveAs();
      } else if (key === "s") {
        event.preventDefault();
        void handleSave();
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [handleOpen, handleSave, handleSaveAs]);

  useEffect(() => {
    let cancelled = false;
    const unlistenPromise = listen<string>("open-file", (event) => {
      if (!univerRef.current || cancelled) {
        return;
      }

      void confirmUnsavedAction()
        .then((canContinue) => {
          if (!canContinue || !univerRef.current || cancelled) {
            return null;
          }
          return runDocumentLoad(() => openPath(univerRef.current!, event.payload));
        })
        .then((state) => {
          if (!state) {
            return;
          }
          stabilizeSpreadsheetFocus();
          return applyDocument(state, `열림: ${state.path ?? "알 수 없는 파일"}`);
        })
        .catch((err) => {
          setError(
            err instanceof Error ? err.message : "파일을 열지 못했습니다.",
          );
        });
    });

    return () => {
      cancelled = true;
      void unlistenPromise.then((unlisten) => unlisten());
    };
  }, [applyDocument, confirmUnsavedAction, runDocumentLoad, stabilizeSpreadsheetFocus]);

  useEffect(() => {
    void syncTitle(doc);
  }, [doc, syncTitle]);

  useEffect(() => {
    let cancelled = false;
    // 앱이 뜨자마자 알림이 튀어나오지 않도록 살짝 지연 (내부적으로 12시간 스로틀도 적용됨)
    const timer = setTimeout(() => {
      void checkForUpdate(APP_VERSION).then((info) => {
        if (!cancelled && info) {
          setUpdateInfo(info);
        }
      });
    }, 2500);

    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, []);

  const handleUpdateLater = useCallback(() => setUpdateInfo(null), []);

  const handleUpdateNow = useCallback(() => {
    if (updateInfo) {
      void openUrl(updateInfo.url);
    }
    setUpdateInfo(null);
  }, [updateInfo]);

  return (
    <div className="app-shell">
      <header className="app-toolbar">
        <div className="app-brand">
          <span className="app-logo" aria-hidden="true">
            <AppLogoIcon />
          </span>
          <div>
            <strong>
              기찬엑셀{" "}
              <span className="app-version" aria-label={`버전 ${APP_VERSION}`}>
                {APP_VERSION}
              </span>
            </strong>
            <p>가볍게 쓰는 스프레드시트</p>
          </div>
        </div>

        <div className="app-actions">
          <button
            type="button"
            className="toolbar-btn"
            title="열기 (Ctrl+O)"
            onClick={() => void handleOpen()}
          >
            <span className="toolbar-emoji" aria-hidden="true">📂</span>
            <span>열기</span>
          </button>
          <button
            type="button"
            className="toolbar-btn"
            title="다른 이름으로 저장 (Ctrl+Shift+S)"
            onClick={() => void handleSaveAs()}
          >
            <span className="toolbar-emoji" aria-hidden="true">📝</span>
            <span>다른 이름으로 저장</span>
          </button>
          <button
            type="button"
            className="toolbar-btn toolbar-btn--primary"
            title="저장 (Ctrl+S)"
            onClick={() => void handleSave()}
          >
            <span className="toolbar-emoji" aria-hidden="true">💾</span>
            <span>저장</span>
          </button>
        </div>
      </header>

      <div className="app-status">
        <span className="app-status-message">
          <span
            className={`app-status-dot${doc.dirty ? " app-status-dot--dirty" : ""}`}
            aria-hidden="true"
          />
          {status}
        </span>
        <span className="app-status-right">
          {error ? <span className="app-error">{error}</span> : null}
          <span className="app-copyright">© 2026 Bae Gichan</span>
        </span>
      </div>

      <main className="app-main" ref={containerRef} tabIndex={-1}>
        <div id="univer-container" className="univer-host" />
      </main>

      {splitDialogOpen ? (
        <div className="dialog-backdrop" role="presentation">
          <section
            aria-labelledby="split-dialog-title"
            className="split-dialog"
            role="dialog"
          >
            <header className="split-dialog-header">
              <h2 id="split-dialog-title">텍스트를 열로 분할</h2>
              <button
                aria-label="닫기"
                type="button"
                onClick={() => setSplitDialogOpen(false)}
              >
                ×
              </button>
            </header>

            <div className="split-dialog-controls">
              <label>
                구분자
                <select
                  value={delimiterMode}
                  onChange={(event) =>
                    setDelimiterMode(event.target.value as DelimiterMode)
                  }
                >
                  <option value="spaces">공백</option>
                  <option value="tab">탭</option>
                  <option value="comma">쉼표</option>
                  <option value="custom">직접 입력</option>
                </select>
              </label>
              {delimiterMode === "custom" ? (
                <label>
                  직접 입력
                  <input
                    autoFocus
                    maxLength={8}
                    type="text"
                    value={customDelimiter}
                    onChange={(event) => setCustomDelimiter(event.target.value)}
                  />
                </label>
              ) : null}
              <label className="split-dialog-check">
                <input
                  checked={mergeDelimiters}
                  type="checkbox"
                  onChange={(event) => setMergeDelimiters(event.target.checked)}
                />
                연속 구분자 묶기
              </label>
            </div>

            <footer className="split-dialog-actions">
              <button type="button" onClick={() => setSplitDialogOpen(false)}>
                취소
              </button>
              <button type="button" onClick={handleSplitSelection}>
                적용
              </button>
            </footer>
          </section>
        </div>
      ) : null}

      {unsavedDialogOpen ? (
        <UnsavedChangesDialog
          fileName={doc.path ? getFileName(doc.path) : null}
          saving={unsavedDialogSaving}
          onSave={() => handleUnsavedChoice("save")}
          onDiscard={() => handleUnsavedChoice("discard")}
          onCancel={() => handleUnsavedChoice("cancel")}
        />
      ) : updateInfo ? (
        <UpdateDialog
          currentVersion={APP_VERSION}
          info={updateInfo}
          onLater={handleUpdateLater}
          onUpdateNow={handleUpdateNow}
        />
      ) : null}
    </div>
  );
}
