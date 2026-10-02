"""Structure tests: core/ must not import streamlit; app exposes its UI entry points."""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CORE_MODULES = [
    "core", "core.config_loader", "core.loader", "core.actions", "core.engine",
    "core.router", "core.merger", "core.exporter", "core.pipeline",
]


def _run(code):
    return subprocess.run(
        [sys.executable, "-c", code], cwd=ROOT, capture_output=True, text=True, check=False,
        env={**__import__("os").environ, "PYTHONIOENCODING": "utf-8"},
    )


def test_core_does_not_import_streamlit():
    code = (
        "import sys\n"
        + "".join(f"import {m}\n" for m in CORE_MODULES)
        + "assert 'streamlit' not in sys.modules, 'streamlit leaked into core'\n"
    )
    res = _run(code)
    assert res.returncode == 0, res.stderr


def test_app_exposes_main_and_load_config_local():
    res = _run("import app; assert callable(app.main); assert hasattr(app, 'load_config_local')")
    assert res.returncode == 0, res.stderr
