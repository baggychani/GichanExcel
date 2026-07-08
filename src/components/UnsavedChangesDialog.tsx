import { useEffect } from "react";
import { WarningIcon } from "./icons";

interface UnsavedChangesDialogProps {
  fileName: string | null;
  saving: boolean;
  onSave: () => void;
  onDiscard: () => void;
  onCancel: () => void;
}

export function UnsavedChangesDialog({
  fileName,
  saving,
  onSave,
  onDiscard,
  onCancel,
}: UnsavedChangesDialogProps) {
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (saving) {
        return;
      }
      if (event.key === "Escape") {
        event.preventDefault();
        onCancel();
      } else if (event.key === "Enter") {
        event.preventDefault();
        onSave();
      }
    };

    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [onCancel, onSave, saving]);

  return (
    <div className="dialog-backdrop" role="presentation">
      <section
        aria-labelledby="unsaved-dialog-title"
        className="confirm-dialog"
        role="alertdialog"
      >
        <div className="confirm-dialog-icon" aria-hidden="true">
          <WarningIcon />
        </div>

        <h2 id="unsaved-dialog-title">저장하지 않은 변경 내용이 있어요</h2>
        <p className="confirm-dialog-subtitle">
          {fileName ? (
            <>
              <strong>{fileName}</strong>에 저장하지 않은 변경 내용이
              있습니다.
            </>
          ) : (
            "계속하기 전에 변경 내용을 저장할까요?"
          )}
        </p>

        <footer className="confirm-dialog-actions">
          <button
            type="button"
            className="confirm-dialog-cancel"
            disabled={saving}
            onClick={onCancel}
          >
            취소
          </button>
          <button
            type="button"
            className="confirm-dialog-discard"
            disabled={saving}
            onClick={onDiscard}
          >
            저장 안 함
          </button>
          <button
            autoFocus
            type="button"
            className="confirm-dialog-save"
            disabled={saving}
            onClick={onSave}
          >
            {saving ? "저장 중…" : "저장"}
          </button>
        </footer>
      </section>
    </div>
  );
}
