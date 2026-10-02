"""Migration planning against MemoryStore (offline)."""
from core.config_loader import read_config_sheets
from store.memory_store import MemoryStore
from store.rules_repo import PRODUCT_ROUTE_FILE, sheets_to_rule_csvs
from tests.regression_harness import CONFIG_PATH
from tools.migrate_rules import apply_plan, build_plan, verify_round_trip

NOW = "2026-10-02T00:00:00+09:00"


def _files() -> dict[str, str]:
    return sheets_to_rule_csvs(read_config_sheets(str(CONFIG_PATH)), "migration", NOW)


def test_round_trip_verification_passes():
    verify_round_trip(str(CONFIG_PATH), _files())


def test_create_then_skip_then_refuse_then_force():
    store, files = MemoryStore(), _files()
    plan = build_plan(store, files, force=False)
    assert {p.action for p in plan} == {"create"}
    apply_plan(store, files, plan)
    assert {p.action for p in build_plan(store, files, force=False)} == {"skip"}

    changed = {**files, PRODUCT_ROUTE_FILE: files[PRODUCT_ROUTE_FILE] + "x"}
    actions = {p.filename: p.action for p in build_plan(store, changed, force=False)}
    assert actions[PRODUCT_ROUTE_FILE] == "refuse"
    forced = build_plan(store, changed, force=True)
    apply_plan(store, changed, forced)
    assert store.read_text(PRODUCT_ROUTE_FILE).content == changed[PRODUCT_ROUTE_FILE]
    assert store.history(PRODUCT_ROUTE_FILE)[0].message == f"chore(data): migrate {PRODUCT_ROUTE_FILE} from config.xlsx"
