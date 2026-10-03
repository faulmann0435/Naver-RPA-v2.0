"""Password gate: pure helpers and the gated app, with AppTest (no network, no real secrets)."""
import pytest
from streamlit.testing.v1 import AppTest

from ui import password_gate
from ui.password_gate import GATE_TITLE, gate_password, password_matches

TIMEOUT = 60
SECRET = "시험-pass"


@pytest.fixture(autouse=True)
def _no_real_secrets(monkeypatch):
    from ui import context

    monkeypatch.setattr(context, "_data_store_settings", lambda: None)


# ---------------------------------------------------------------- pure helpers

@pytest.mark.parametrize(
    ("secrets_like", "expected"),
    [
        ({}, None),
        ({"app_gate": {}}, None),
        ({"app_gate": {"password": ""}}, None),
        ({"app_gate": {"password": "   "}}, None),
        ({"app_gate": "oops"}, None),
        ({"app_gate": {"password": "abc"}}, "abc"),
        ({"app_gate": {"password": 1234}}, "1234"),
    ],
)
def test_gate_password(secrets_like, expected):
    assert gate_password(secrets_like) == expected


def test_gate_password_survives_broken_secrets():
    class Broken:
        def __contains__(self, _key):
            raise FileNotFoundError("no secrets.toml")

    assert gate_password(Broken()) is None


@pytest.mark.parametrize(
    ("entered", "expected", "ok"),
    [
        ("abc", "abc", True),
        ("  abc \n", "abc", True),
        ("abd", "abc", False),
        ("ABC", "abc", False),
        ("", "abc", False),
        ("abcd", "abc", False),
        (SECRET, SECRET, True),
        ("시험-pas", SECRET, False),
    ],
)
def test_password_matches(entered, expected, ok):
    assert password_matches(entered, expected) is ok


# ---------------------------------------------------------------- gated app

def _app(secret: str | None = SECRET) -> AppTest:
    at = AppTest.from_file("app.py", default_timeout=TIMEOUT)
    if secret is not None:
        at.secrets["app_gate"] = {"password": secret}
    return at.run()


def _submit(at: AppTest, text: str) -> AppTest:
    at.text_input[0].set_value(text)
    at.button[0].click()
    return at.run()


def _errors(at: AppTest) -> list[str]:
    return [e.value for e in at.error]


def test_no_secret_means_no_gate():
    at = _app(secret=None)
    assert not at.exception
    assert not at.text_input or all(t.label != "비밀번호" for t in at.text_input)
    assert GATE_TITLE not in [t.value for t in at.title]
    assert at.sidebar.selectbox  # the normal app (worker picker) rendered


def test_empty_secret_means_no_gate():
    at = _app(secret="")
    assert GATE_TITLE not in [t.value for t in at.title]
    assert at.sidebar.selectbox


def test_secret_shows_gate_and_no_pages():
    at = _app()
    assert not at.exception
    assert [t.value for t in at.title] == [GATE_TITLE]
    assert at.text_input[0].label == "비밀번호"
    assert at.text_input[0].proto.type == 1  # password input
    assert at.button[0].label == "들어가기"
    assert not at.sidebar.selectbox  # no sidebar data access before the password


def test_wrong_password_shows_error():
    at = _submit(_app(), "nope")
    assert "비밀번호가 맞지 않습니다." in _errors(at)
    assert [t.value for t in at.title] == [GATE_TITLE]
    assert not at.sidebar.selectbox


def test_correct_password_renders_app_and_keeps_no_copy():
    at = _submit(_app(), f"  {SECRET} ")
    assert not at.exception
    assert GATE_TITLE not in [t.value for t in at.title]
    assert at.sidebar.selectbox
    assert SECRET not in repr({k: v for k, v in at.session_state.filtered_state.items()})


def test_lockout_after_five_failures_then_release(monkeypatch):
    clock = {"t": 1000.0}
    monkeypatch.setattr(password_gate, "_now", lambda: clock["t"])
    at = _app()
    for _ in range(4):
        at = _submit(at, "bad")
        assert "비밀번호가 맞지 않습니다." in _errors(at)
    at = _submit(at, "bad")  # 5th failure starts the lockout
    assert any("60초" in e for e in _errors(at))
    # Even the correct password is refused while locked.
    at = _submit(at, SECRET)
    assert any("초 뒤에 다시 시도" in e for e in _errors(at))
    assert not at.sidebar.selectbox
    clock["t"] += 61
    at = _submit(at, SECRET)
    assert at.sidebar.selectbox
