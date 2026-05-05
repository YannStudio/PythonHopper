import json

import pandas as pd

from export_session_log import (
    EXPORT_SESSION_LOG_FILENAME,
    build_export_session_log,
    convert_offers_to_orders,
    find_export_session_logs,
    load_export_session_log,
    write_export_session_log,
)
from orders import _apply_order_pricing


def test_export_session_log_roundtrip(tmp_path):
    state = {
        "selections": {"production::Cutting": "MCB"},
        "doc_types": {"production::Cutting": "Offerteaanvraag"},
        "doc_numbers": {"production::Cutting": "OFF-100"},
        "pricing": {
            "production::Cutting": {
                "unit_price": "12.50",
                "total_price": "",
            }
        },
        "remember": False,
    }
    bom = pd.DataFrame(
        [
            {
                "PartNumber": "A",
                "Production": "Cutting",
                "Aantal": 2,
            }
        ]
    )

    payload = build_export_session_log(
        project_number="20250165",
        project_name="Piva",
        client_name="Tecno Art bvba",
        bom_source_path="C:/tmp/bom.xlsx",
        bom_df=bom,
        state=state,
        app_version="test",
    )
    path = write_export_session_log(tmp_path, payload)

    assert path.endswith(EXPORT_SESSION_LOG_FILENAME)
    loaded = load_export_session_log(path)
    assert loaded["project"]["number"] == "20250165"
    assert loaded["order_state"]["selections"]["production::Cutting"] == "MCB"
    assert loaded["order_state"]["pricing"]["production::Cutting"]["unit_price"] == "12.50"
    assert loaded["bom"]["row_count"] == 1
    assert loaded["bom"]["sha256"]


def test_convert_offers_to_orders_clears_off_numbers():
    converted = convert_offers_to_orders(
        {
            "doc_types": {
                "production::Cutting": "Offerteaanvraag",
                "production::Roof": "Bestelbon",
            },
            "doc_numbers": {
                "production::Cutting": "OFF-123",
                "production::Roof": "BB-456",
            },
        }
    )

    assert converted["doc_types"]["production::Cutting"] == "Bestelbon"
    assert converted["doc_numbers"]["production::Cutting"] == ""
    assert converted["doc_numbers"]["production::Roof"] == "BB-456"


def test_export_session_log_rejects_unknown_schema(tmp_path):
    path = tmp_path / EXPORT_SESSION_LOG_FILENAME
    path.write_text(json.dumps({"schema_version": 999}), encoding="utf-8")

    try:
        load_export_session_log(path)
    except ValueError as exc:
        assert "Niet-ondersteunde" in str(exc)
    else:
        raise AssertionError("Expected ValueError for unsupported schema")


def test_find_export_session_logs_newest_first(tmp_path):
    older_dir = tmp_path / "2026-01-01_project"
    newer_dir = tmp_path / "2026-01-02_project"
    older_dir.mkdir()
    newer_dir.mkdir()
    older = older_dir / EXPORT_SESSION_LOG_FILENAME
    newer = newer_dir / EXPORT_SESSION_LOG_FILENAME
    older.write_text("{}", encoding="utf-8")
    newer.write_text("{}", encoding="utf-8")
    older.touch()
    newer.touch()

    logs = find_export_session_logs(tmp_path)

    assert logs[0] == str(newer)
    assert str(older) in logs


def test_apply_order_pricing_adds_unit_and_total_columns():
    items, layout = _apply_order_pricing(
        [{"PartNumber": "A", "Description": "Plaat", "Aantal": 3}],
        {"unit_price": "12,50", "total_price": "40"},
        context_kind="Productie",
    )

    assert layout is not None
    assert items[0]["Eenheidsprijs"] == "12.50"
    assert items[0]["Totaalprijs"] == "37.50"
    assert items[-1]["Description"] == "Totaal aangeboden"
    assert items[-1]["Totaalprijs"] == "40"
