"""Keep every test away from the real 자연이랑 files under DATA_DIR/발주기록."""
import pytest


@pytest.fixture(autouse=True)
def _isolated_withnature_files(tmp_path, monkeypatch):
    monkeypatch.setattr("worker.withnature_run.ORDERS_FILE", tmp_path / "wn_orders.json")
    monkeypatch.setattr("worker.invoice_run.WITHNATURE_LEDGER_FILE", tmp_path / "wn_ledger.json")
