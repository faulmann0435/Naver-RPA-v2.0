"""get_config: data-store failures must fall back to config.xlsx instead of stopping the app."""
import warnings

import pytest
import requests

warnings.filterwarnings("ignore")

import app
from store.base import StoreError
from ui import context


@pytest.fixture
def configured(monkeypatch):
    monkeypatch.setattr(context, "_data_store_settings", lambda: ("owner/data", "dev"))
    monkeypatch.setattr(context, "load_config_local", lambda *a, **k: {"source": "xlsx"})


@pytest.mark.parametrize(
    "error",
    [
        StoreError("GitHub 연결 실패"),
        requests.ConnectionError("down"),
        KeyError("token"),
        ValueError("ProductRoute에 'Priority' 컬럼이 없습니다."),
    ],
)
def test_falls_back_to_xlsx_on_store_failure(configured, monkeypatch, error):
    def boom(*_args, **_kwargs):
        raise error

    monkeypatch.setattr(context, "_load_config_from_data_store", boom)
    config, label = app.get_config("config.xlsx", "key")
    assert config == {"source": "xlsx"}
    assert label == "config.xlsx"


def test_uses_store_when_available(configured, monkeypatch):
    monkeypatch.setattr(context, "_load_config_from_data_store", lambda *a: {"source": "store"})
    config, label = app.get_config("config.xlsx", "key")
    assert config == {"source": "store"}
    assert label == "데이터 저장소 (owner/data@dev)"


def test_uses_xlsx_without_secrets(monkeypatch):
    monkeypatch.setattr(context, "_data_store_settings", lambda: None)
    monkeypatch.setattr(context, "load_config_local", lambda *a, **k: {"source": "xlsx"})
    assert app.get_config("config.xlsx", "key") == ({"source": "xlsx"}, "config.xlsx")
