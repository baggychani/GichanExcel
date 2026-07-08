import { getCurrentWindow } from "@tauri-apps/api/window";
import { listen } from "@tauri-apps/api/event";
import { confirm } from "@tauri-apps/plugin-dialog";
import { openUrl } from "@tauri-apps/plugin-opener";
import { useCallback, useEffect, useRef, useState } from "react";
import type { FUniver } from "@univerjs/core/facade";
import {
  getFileName,
  getWindowTitle,
  loadWorkbookData,
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
  clearAutoSave,
  readAutoSave,
  writeAutoSave,
} from "../lib/autosave";
import { checkForUpdate, type UpdateInfo } from "../lib/update-checker";
import { TEXT_SPLIT_EVENT } from "../plugins/text-split";
import { setupUniver } from "../setup-univer";
import { APP_VERSION } from "../lib/version";
import { UpdateDialog } from "./UpdateDialog";
import { UnsavedChangesDialog } from "./UnsavedChangesDialog";
import { AppLogoIcon, FolderOpenIcon, SaveAsIcon, SaveIcon } from "./icons";

type UnsavedChoice = "save" | "discard" | "cancel";

const INITIAL_DOC: DocumentState = { path: null, dirty: false };

export function SpreadsheetApp() {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<FUniver | null>(null);
  const docRef = useRef<DocumentState>(INITIAL_DOC);
  const closeAllowedRef = useRef(false);
  const unsavedChoiceResolverRef = useRef<((choice: UnsavedChoice) => void) | null>(
    null,
  );
  const autoFittingRowsRef = useRef(false);
  const [doc, setDoc] = useState<DocumentState>(INITIAL_DOC);
  const [ready, setReady] = useState(false);
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

  const markDirty = useCallback(() => {
    setDoc((current) => {
      if (current.dirty) {
        return current;
      }
      const next = { ...current, dirty: true };
      void syncTitle(next);
      return next;
    });
    setAutoSaveVersion((version) => version + 1);
  }, [syncTitle]);

  const applyDocument = useCallback(
    async (state: DocumentState, message: string) => {
      setDoc(state);
      setStatus(message);
      setError(null);
      await syncTitle(state);
      if (!state.dirty) {
        void clearAutoSave();
      }
    },
    [syncTitle],
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

      const state = await openSpreadsheet(univerRef.current);
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
  }, [applyDocument, confirmUnsavedAction]);

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

  useEffect(() => {
    if (!containerRef.current) {
      return;
    }

    const app = setupUniver("univer-container");
    univerRef.current = app.univerAPI;
    setReady(true);

    return () => {
      setReady(false);
      app.dispose();
      univerRef.current = null;
    };
  }, []);

  useEffect(() => {
    if (!ready || !univerRef.current) {
      return;
    }

    let cancelled = false;
    void readAutoSave()
      .then(async (record) => {
        if (!record || cancelled || !univerRef.current) {
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

        loadWorkbookData(univerRef.current, record.snapshot);
        await applyDocument(
          { path: record.path, dirty: true },
          `자동저장 복구본을 열었습니다. (${savedAt})`,
        );
      })
      .catch(() => {
        setError("자동저장 복구본을 읽지 못했습니다.");
      });

    return () => {
      cancelled = true;
    };
  }, [applyDocument, ready]);

  useEffect(() => {
    const onOpenTextSplit = () => openSplitDialog();
    window.addEventListener(TEXT_SPLIT_EVENT, onOpenTextSplit);
    return () => window.removeEventListener(TEXT_SPLIT_EVENT, onOpenTextSplit);
  }, [openSplitDialog]);

  useEffect(() => {
    const univerAPI = univerRef.current;
    if (!univerAPI) {
      return;
    }

    const disposable = univerAPI.addEvent(
      univerAPI.Event.CommandExecuted,
      () => {
        markDirty();
        if (autoFittingRowsRef.current) {
          return;
        }

        const workbook = univerAPI.getActiveWorkbook();
        const sheet = workbook?.getActiveSheet();
        const activeRange = workbook?.getActiveRange();
        if (sheet && activeRange) {
          autoFittingRowsRef.current = true;
          try {
            sheet.setRangesAutoHeight([activeRange.getRange()]);
          } finally {
            autoFittingRowsRef.current = false;
          }
        }
      },
    );

    return () => disposable.dispose();
  }, [markDirty]);

  useEffect(() => {
    const window = getCurrentWindow();
    const unlistenPromise = window.onCloseRequested(async (event) => {
      if (closeAllowedRef.current || !docRef.current.dirty) {
        return;
      }

      event.preventDefault();
      if (unsavedChoiceResolverRef.current) {
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

    const timer = setTimeout(() => {
      const workbook = univerRef.current?.getActiveWorkbook();
      if (!workbook || !docRef.current.dirty) {
        return;
      }

      void writeAutoSave({
        path: docRef.current.path,
        savedAt: new Date().toISOString(),
        snapshot: workbook.save(),
      })
        .then(() => {
          setStatus(`자동저장됨 ${new Date().toLocaleTimeString()}`);
        })
        .catch(() => {
          setError("자동저장을 하지 못했습니다.");
        });
    }, 4000);

    return () => clearTimeout(timer);
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
          return openPath(univerRef.current, event.payload);
        })
        .then((state) => {
          if (!state) {
            return;
          }
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
  }, [applyDocument, confirmUnsavedAction]);

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
            <FolderOpenIcon />
            <span>열기</span>
          </button>
          <button
            type="button"
            className="toolbar-btn"
            title="다른 이름으로 저장 (Ctrl+Shift+S)"
            onClick={() => void handleSaveAs()}
          >
            <SaveAsIcon />
            <span>다른 이름으로 저장</span>
          </button>
          <button
            type="button"
            className="toolbar-btn toolbar-btn--primary"
            title="저장 (Ctrl+S)"
            onClick={() => void handleSave()}
          >
            <SaveIcon />
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

      <main className="app-main" ref={containerRef}>
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
