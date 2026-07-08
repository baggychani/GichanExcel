import { getCurrentWindow } from "@tauri-apps/api/window";
import { listen } from "@tauri-apps/api/event";
import { useCallback, useEffect, useRef, useState } from "react";
import type { FUniver } from "@univerjs/core/facade";
import {
  getWindowTitle,
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
import { TEXT_SPLIT_EVENT } from "../plugins/text-split";
import { setupUniver } from "../setup-univer";
import { APP_VERSION } from "../lib/version";
import { AppLogoIcon, FolderOpenIcon, SaveAsIcon, SaveIcon } from "./icons";

const INITIAL_DOC: DocumentState = { path: null, dirty: false };

export function SpreadsheetApp() {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<FUniver | null>(null);
  const [doc, setDoc] = useState<DocumentState>(INITIAL_DOC);
  const [status, setStatus] = useState("새 통합 문서");
  const [error, setError] = useState<string | null>(null);
  const [delimiterMode, setDelimiterMode] = useState<DelimiterMode>("spaces");
  const [customDelimiter, setCustomDelimiter] = useState("");
  const [mergeDelimiters, setMergeDelimiters] = useState(true);
  const [splitDialogOpen, setSplitDialogOpen] = useState(false);

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
  }, [syncTitle]);

  const applyDocument = useCallback(
    async (state: DocumentState, message: string) => {
      setDoc(state);
      setStatus(message);
      setError(null);
      await syncTitle(state);
    },
    [syncTitle],
  );

  const handleOpen = useCallback(async () => {
    if (!univerRef.current) {
      return;
    }

    try {
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
  }, [applyDocument]);

  const handleSave = useCallback(async () => {
    if (!univerRef.current) {
      return;
    }

    try {
      const state = await saveSpreadsheet(univerRef.current, doc.path);
      await applyDocument(state, "저장했습니다.");
    } catch (err) {
      if (err instanceof Error && err.message === "cancelled") {
        return;
      }
      setError(err instanceof Error ? err.message : "저장하지 못했습니다.");
    }
  }, [applyDocument, doc.path]);

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

    return () => {
      app.dispose();
      univerRef.current = null;
    };
  }, []);

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
      () => markDirty(),
    );

    return () => disposable.dispose();
  }, [markDirty]);

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

      void openPath(univerRef.current, event.payload)
        .then((state) =>
          applyDocument(state, `열림: ${state.path ?? "알 수 없는 파일"}`),
        )
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
  }, [applyDocument]);

  useEffect(() => {
    void syncTitle(doc);
  }, [doc, syncTitle]);

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
        {error ? <span className="app-error">{error}</span> : null}
      </div>

      <main className="app-main" ref={containerRef}>
        <div id="univer-container" className="univer-host" />
      </main>

      {splitDialogOpen ? (
        <div className="split-dialog-backdrop" role="presentation">
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
    </div>
  );
}
