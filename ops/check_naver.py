"""Setup check: can this PC get a Naver Commerce API token? Prints OK or the error; exit code 0 / 2."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from worker.naver_client import fetch_token, load_credentials  # noqa: E402

try:
    fetch_token(load_credentials())
except Exception as exc:  # setup tool: show whatever went wrong
    print(f"{type(exc).__name__}: {exc}")
    raise SystemExit(2)
print("OK")
