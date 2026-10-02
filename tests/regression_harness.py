"""
Golden-master regression harness.

Runs the order pipeline on real Naver order files (kept OUTSIDE the repo, they contain
customer data) and stores / compares the generated purchase-order workbooks cell by cell.

Usage:
    python -m tests.regression_harness --update     # (re)build tests/golden from current code
    python -m pytest tests/test_regression.py       # compare current code against golden

Sample directory: env RPA_SAMPLE_DIR, default = parent folder of the repo.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import warnings
from collections.abc import Callable
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parent.parent
GOLDEN_DIR = REPO_ROOT / "tests" / "golden"
CONFIG_PATH = REPO_ROOT / "config.xlsx"
REQUIRED_COLUMNS = {"상품명", "옵션정보", "수량", "배송비 묶음번호"}
_DATE_SUFFIX = re.compile(r"_\d{8}(?=\.xlsx$)")

if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

warnings.filterwarnings("ignore")


def sample_dir() -> Path:
    return Path(os.environ.get("RPA_SAMPLE_DIR", REPO_ROOT.parent)).resolve()


class _Upload:
    """Mimics Streamlit's UploadedFile (name + read())."""

    def __init__(self, path: Path):
        self.name = path.name
        self._data = path.read_bytes()

    def read(self) -> bytes:
        return self._data


def _entrypoints():
    """Return (load_config, load_order, process_all_data) from the refactored package or legacy app."""
    try:
        from core.config_loader import load_config
        from core.loader import load_excel
        from core.pipeline import process_all_data
    except ImportError:
        import app

        load_config = app.load_config_local.__wrapped__
        load_excel = app.load_excel
        process_all_data = app.process_all_data
    return load_config, load_excel, process_all_data


def discover_inputs() -> list[Path]:
    """All readable (unencrypted) order files with the required columns, outside the repo."""
    _, load_excel, _ = _entrypoints()
    found = []
    for path in sorted(sample_dir().rglob("*.xlsx")):
        if REPO_ROOT in path.parents or path.name.startswith("~$") or path.name == "config.xlsx":
            continue
        try:
            df = load_excel(_Upload(path))
        except Exception:
            continue  # encrypted, not an order file, etc.
        if REQUIRED_COLUMNS <= set(df.columns):
            found.append(path)
    return found


def case_key(path: Path) -> str:
    rel = path.relative_to(sample_dir()).as_posix()
    return f"{path.stem}__{hashlib.sha1(rel.encode('utf-8')).hexdigest()[:8]}"


@dataclass(frozen=True)
class OutputFile:
    name: str          # filename without the _YYYYMMDD date suffix
    vendor: str
    data: bytes


def run_case(path: Path, config: dict | None = None, runner: Callable | None = None) -> list[OutputFile]:
    """Run the pipeline; uses `config` when given, else loads config.xlsx.

    `runner(df, config) -> list[dict]` replaces process_all_data when given.
    """
    load_config, load_excel, process_all_data = _entrypoints()
    if config is None:
        config = load_config(str(CONFIG_PATH))
    df = load_excel(_Upload(path))
    results = (runner or process_all_data)(df, config)
    return [
        OutputFile(name=_DATE_SUFFIX.sub("", r["filename"]), vendor=r["vendor"], data=r["data"].getvalue())
        for r in results
    ]


def _cells(data: bytes) -> pd.DataFrame:
    return pd.read_excel(BytesIO(data), dtype=object, keep_default_na=False)


def _conditional_formats(data: bytes) -> list[str]:
    ws = load_workbook(BytesIO(data)).active
    return sorted(str(rng.sqref) for rng in ws.conditional_formatting)


def update_golden() -> None:
    inputs = discover_inputs()
    if not inputs:
        raise SystemExit(f"No readable order files under {sample_dir()}")
    GOLDEN_DIR.mkdir(parents=True, exist_ok=True)
    manifest = {}
    for path in inputs:
        key = case_key(path)
        case_dir = GOLDEN_DIR / key
        case_dir.mkdir(parents=True, exist_ok=True)
        for old in case_dir.glob("*.xlsx"):
            old.unlink()
        outputs = run_case(path)
        for out in outputs:
            (case_dir / out.name).write_bytes(out.data)
        manifest[key] = {
            "input": path.relative_to(sample_dir()).as_posix(),
            "files": {o.name: {"vendor": o.vendor, "rows": len(_cells(o.data))} for o in outputs},
        }
        print(f"[golden] {key}: {len(outputs)} files")
    (GOLDEN_DIR / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[golden] {len(manifest)} cases written to {GOLDEN_DIR}")


def load_manifest() -> dict:
    path = GOLDEN_DIR / "manifest.json"
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def compare_case(key: str, entry: dict, config: dict | None = None, runner: Callable | None = None) -> list[str]:
    """Return a list of human-readable differences (empty = identical)."""
    path = sample_dir() / entry["input"]
    if not path.exists():
        return [f"input missing: {path}"]
    actual = {o.name: o for o in run_case(path, config, runner)}
    expected = set(entry["files"])
    problems = []
    if set(actual) != expected:
        problems.append(f"file set differs: expected {sorted(expected)}, got {sorted(actual)}")
    for name in sorted(expected & set(actual)):
        golden_bytes = (GOLDEN_DIR / key / name).read_bytes()
        out = actual[name]
        if out.vendor != entry["files"][name]["vendor"]:
            problems.append(f"{name}: vendor {out.vendor!r} != {entry['files'][name]['vendor']!r}")
        exp_df, act_df = _cells(golden_bytes), _cells(out.data)
        if list(exp_df.columns) != list(act_df.columns):
            problems.append(f"{name}: columns differ {list(exp_df.columns)} vs {list(act_df.columns)}")
        elif exp_df.shape != act_df.shape:
            problems.append(f"{name}: shape {exp_df.shape} vs {act_df.shape}")
        else:
            diff = exp_df.astype(str).ne(act_df.astype(str))
            for r, c in zip(*diff.to_numpy().nonzero()):
                problems.append(
                    f"{name} row {r + 2} col {exp_df.columns[c]!r}: "
                    f"{exp_df.iat[r, c]!r} -> {act_df.iat[r, c]!r}"
                )
        if _conditional_formats(golden_bytes) != _conditional_formats(out.data):
            problems.append(f"{name}: conditional formatting (제주 highlight) differs")
    return problems


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--update", action="store_true", help="rebuild golden files from current code")
    args = parser.parse_args()
    if args.update:
        update_golden()
        return
    failures = {k: compare_case(k, v) for k, v in load_manifest().items()}
    for key, problems in failures.items():
        print(f"{'OK  ' if not problems else 'FAIL'} {key}")
        for p in problems[:20]:
            print("     ", p)
    if any(failures.values()):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
