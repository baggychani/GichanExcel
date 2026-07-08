import type { UpdateInfo } from "../lib/update-checker";
import { SparkleIcon } from "./icons";

interface UpdateDialogProps {
  currentVersion: string;
  info: UpdateInfo;
  onLater: () => void;
  onUpdateNow: () => void;
}

export function UpdateDialog({
  currentVersion,
  info,
  onLater,
  onUpdateNow,
}: UpdateDialogProps) {
  return (
    <div className="dialog-backdrop" role="presentation">
      <section
        aria-labelledby="update-dialog-title"
        className="update-dialog"
        role="dialog"
      >
        <div className="update-dialog-icon" aria-hidden="true">
          <SparkleIcon />
        </div>

        <h2 id="update-dialog-title">새 버전이 나왔어요</h2>
        <p className="update-dialog-subtitle">
          GitHub Releases에서 최신 설치 파일을 받을 수 있어요.
        </p>

        <div className="update-dialog-versions">
          <span className="version-pill version-pill--current">
            현재 {currentVersion}
          </span>
          <span className="update-dialog-arrow" aria-hidden="true">
            →
          </span>
          <span className="version-pill version-pill--latest">
            최신 {info.version}
          </span>
        </div>

        {info.notes ? (
          <div className="update-dialog-notes">
            <p className="update-dialog-notes-label">릴리스 노트</p>
            <div className="update-dialog-notes-body">{info.notes}</div>
          </div>
        ) : null}

        <footer className="update-dialog-actions">
          <button type="button" onClick={onLater}>
            나중에
          </button>
          <button
            type="button"
            className="update-dialog-actions-primary"
            onClick={onUpdateNow}
          >
            지금 업데이트
          </button>
        </footer>
      </section>
    </div>
  );
}
