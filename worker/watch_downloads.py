"""Watch the Downloads folder and line every new Naver order export up against the API.

    python -m worker.watch_downloads            # keep watching (Ctrl+C to stop)
    python -m worker.watch_downloads --once     # check once and exit

On its first start it only remembers the exports already there; it compares exports that
arrive afterwards, once each, as soon as the download has finished. Each result is one line
on screen (later: a Telegram message) and a full masked report under 자동대조/ next to the repo.
Encrypted exports need RPA_ORDER_PASSWORD (environment or .env).
The folder is RPA_DOWNLOAD_FOLDER in .env (default: the user's Downloads); --folder overrides it.
Pointing it at another folder starts over: exports already in the new folder are skipped.
"""
from __future__ import annotations

import argparse
import io
import json
import logging
import os
import sys
import time
from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Protocol

from worker.compare import CompareResult, compare_file
from worker.naver_adapter import KST
from worker.naver_client import ENV_FILE, read_env_file
from worker.paths import DATA_DIR

DEFAULT_FOLDER = Path.home() / "Downloads"
PATTERN = "스마트스토어_*.xlsx"
POLL_SECONDS = 10
SETTLE_SECONDS = 5  # a file untouched this long has finished downloading
OUT_DIR = DATA_DIR / "자동대조"
STATE_FILE = OUT_DIR / "처리기록.json"
PASSWORD_KEY = "RPA_ORDER_PASSWORD"
FOLDER_KEY = "RPA_DOWNLOAD_FOLDER"

log = logging.getLogger("watch_downloads")


class Notifier(Protocol):
    def send(self, text: str) -> None: ...


class ConsoleNotifier:
    def send(self, text: str) -> None:
        print(text, flush=True)


def _signature(path: Path) -> str:
    stat = path.stat()
    return f"{stat.st_size}:{int(stat.st_mtime)}"


def load_state(folder: Path, path: Path = STATE_FILE) -> dict[str, str] | None:
    """Exports already handled in `folder` (name -> size:mtime); None on a first run or a new folder."""
    if not path.exists():
        return None
    saved = json.loads(path.read_text(encoding="utf-8"))
    return saved.get("files", {}) if saved.get("folder") == str(folder.resolve()) else None


def save_state(folder: Path, files: dict[str, str], path: Path = STATE_FILE) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    body = {"folder": str(folder.resolve()), "files": files}
    path.write_text(json.dumps(body, ensure_ascii=False, indent=2), encoding="utf-8")


def configured_folder() -> Path:
    """RPA_DOWNLOAD_FOLDER from the environment or .env, else the user's Downloads folder."""
    value = os.environ.get(FOLDER_KEY) or read_env_file(ENV_FILE).get(FOLDER_KEY)
    return Path(value).expanduser() if value else DEFAULT_FOLDER


def finished_exports(folder: Path, now: float) -> list[Path]:
    """Order exports in the folder that are not being written any more (oldest first)."""
    files = [p for p in folder.glob(PATTERN) if p.is_file() and p.stat().st_size > 0]
    settled = [p for p in files if now - p.stat().st_mtime >= SETTLE_SECONDS]
    return sorted(settled, key=lambda p: p.stat().st_mtime)


def new_exports(folder: Path, state: dict[str, str], now: float) -> list[Path]:
    """Finished exports not handled yet (a re-download under the same name counts as new)."""
    return [p for p in finished_exports(folder, now) if state.get(p.name) != _signature(p)]


def write_report(result: CompareResult, when: datetime, out_dir: Path = OUT_DIR) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    report = out_dir / f"{when:%Y%m%d_%H%M%S}_{Path(result.file).stem}.txt"
    report.write_text("\n".join([result.summary(), "", *result.details()]) + "\n", encoding="utf-8")
    return report


def _safe_compare(path: Path, compare: Callable[[Path], CompareResult]) -> CompareResult:
    try:
        return compare(path)
    except Exception as exc:  # one bad file must not stop the watcher
        log.exception("대조 실패: %s", path.name)
        return CompareResult(path.name, error=f"{type(exc).__name__}: {exc}")


def check_once(folder: Path, notifier: Notifier, compare: Callable[[Path], CompareResult] = compare_file,
               state_file: Path = STATE_FILE, out_dir: Path = OUT_DIR) -> list[CompareResult]:
    """Compare the exports that arrived since the last check; the first check only takes stock."""
    now = time.time()
    state = load_state(folder, state_file)
    if state is None:
        baseline = {p.name: _signature(p) for p in finished_exports(folder, now)}
        save_state(folder, baseline, state_file)
        notifier.send(f"자동 대조 시작: {folder} 감시 중 (이미 있던 주문 엑셀 {len(baseline)}개는 건너뜀)")
        return []
    results = []
    for path in new_exports(folder, state, now):
        result = _safe_compare(path, compare)
        write_report(result, datetime.now(KST), out_dir)
        notifier.send(result.summary())
        state = {**state, path.name: _signature(path)}
        save_state(folder, state, state_file)
        results.append(result)
    return results


def _load_password() -> None:
    if not os.environ.get(PASSWORD_KEY):
        value = read_env_file(ENV_FILE).get(PASSWORD_KEY)
        if value:
            os.environ[PASSWORD_KEY] = value


def main(argv: list[str] | None = None) -> int:
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--folder", type=Path, default=None, help=f"default: {FOLDER_KEY} in .env, else Downloads")
    parser.add_argument("--once", action="store_true", help="check once and exit")
    args = parser.parse_args(argv)
    folder = args.folder or configured_folder()
    if not folder.is_dir():
        print(f"폴더가 없습니다: {folder} (.env의 {FOLDER_KEY} 확인)")
        return 2
    _load_password()
    notifier = ConsoleNotifier()
    check_once(folder, notifier)
    while not args.once:
        time.sleep(POLL_SECONDS)
        check_once(folder, notifier)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
