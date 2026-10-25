"""The owner's PC installs only requirements-bot.txt: it must hold every package except the web UI."""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WEB_ONLY = {"streamlit"}


def _lines(name: str) -> set[str]:
    text = (ROOT / name).read_text(encoding="utf-8")
    return {line.strip() for line in text.splitlines() if line.strip() and not line.startswith("#")}


def test_bot_requirements_match_app_requirements_minus_web_ui():
    app = {line for line in _lines("requirements.txt") if line.split("=")[0].split(">")[0] not in WEB_ONLY}
    assert _lines("requirements-bot.txt") == app
