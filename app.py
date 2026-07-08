from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent


def pause() -> None:
    if os.name == "nt":
        input("\nPress Enter to close this window...")


def main() -> int:
    os.chdir(ROOT)

    npm = shutil.which("npm.cmd") or shutil.which("npm")
    if not npm:
        print("Node.js/npm was not found.")
        print("Install Node.js LTS, then run this launcher again.")
        return 1

    print("Starting GichanExcel in development mode...")
    print("This window must stay open while the app is running.\n")
    return subprocess.call([npm, "run", "tauri", "dev"])


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)
    except Exception as exc:
        print(f"Failed to start GichanExcel: {exc}", file=sys.stderr)
        pause()
        raise SystemExit(1)
