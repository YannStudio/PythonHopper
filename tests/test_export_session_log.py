import json

import pandas as pd

from export_session_log import (
    EXPORT_SESSION_LOG_FILENAME,
    build_export_session_log,
    convert_offers_to_orders,
    find_export_session_logs,
    format_export_log_compatibility_message,
    load_export_session_log,
    merge_order_state_sections,
    normalize_spare_parts_info,
    resolve_export_document_path,
    state_keys_for_import_sections,
    summarize_export_log_compatibility,
    write_export_session_log,
)
from orders import _apply_order_pricing


def test_export_session_log_roundtrip(tmp_path):
    state = {
        "selections": {"production::Cutting": "MCB"},
        "doc_types": {"production::Cutting": "Offerteaanvraag"},
        "doc_numbers": {"production::Cutting": "OFF-100"},
        "vat_rates": {"production::Cutting": "21"},
        "pricing": {
            "production::Cutting": {
                "unit_price": "12.50",
                "total_price": "",
                "items": {
                    "part|A": {
                        "unit_price": "12.50",
                        "total_price": "25.00",
                    }
                },
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
        generated_documents=[
            {
                "path": "Cutting/Bestelbon_BB-100_Cutting.pdf",
                "kind": "order",
                "format": "pdf",
                "selection_key": "production::Cutting",
                "selection_keys": ["production::Cutting"],
                "doc_type": "Bestelbon",
                "doc_number": "BB-100",
                "supplier": "MCB",
            },
            {"path": ""},
        ],
        status_messages=["Bestelbon opgeslagen.", "Bestelbon opgeslagen.", ""],
        path_limit_warnings=["pad ingekort"],
        spare_parts={
            "group_overrides": {"sparepart:1|PN1": "Electro"},
            "groups": [
                {
                    "key": "custom--electro",
                    "label": "Electro",
                    "display_label": "Spare Parts - Electro",
                    "document_label": "Spare Parts - Electro",
                    "route_source": "custom",
                    "item_count": "2",
                    "missing_count": 1,
                }
            ],
        },
    )
    path = write_export_session_log(tmp_path, payload)

    assert path.endswith(EXPORT_SESSION_LOG_FILENAME)
    loaded = load_export_session_log(path)
    assert loaded["project"]["number"] == "20250165"
    assert loaded["order_state"]["selections"]["production::Cutting"] == "MCB"
    assert loaded["order_state"]["pricing"]["production::Cutting"]["unit_price"] == "12.50"
    assert loaded["order_state"]["vat_rates"]["production::Cutting"] == "21"
    assert (
        loaded["order_state"]["pricing"]["production::Cutting"]["items"]["part|A"][
            "total_price"
        ]
        == "25.00"
    )
    assert loaded["bom"]["row_count"] == 1
    assert loaded["bom"]["sha256"]
    assert loaded["export"]["generated_documents"] == [
        {
            "path": "Cutting/Bestelbon_BB-100_Cutting.pdf",
            "kind": "order",
            "format": "pdf",
            "selection_key": "production::Cutting",
            "doc_type": "Bestelbon",
            "doc_number": "BB-100",
            "supplier": "MCB",
            "selection_keys": ["production::Cutting"],
        }
    ]
    assert loaded["export"]["status_messages"] == ["Bestelbon opgeslagen."]
    assert loaded["export"]["path_limit_warnings"] == ["pad ingekort"]
    assert loaded["spare_parts"]["group_overrides"] == {"sparepart:1|PN1": "Electro"}
    assert loaded["spare_parts"]["groups"][0]["key"] == "custom--electro"
    assert loaded["spare_parts"]["groups"][0]["item_count"] == 2


def test_normalize_spare_parts_info_keeps_overrides_and_group_summaries():
    info = normalize_spare_parts_info(
        {
            "group_overrides": {
                "sparepart:1|PN1": "Electro",
                "": "Ignored",
                "sparepart:2|PN2": " ",
            },
            "groups": [
                {
                    "key": "custom--electro",
                    "label": "Electro",
                    "display_label": "Spare Parts - Electro",
                    "document_label": "Spare Parts - Electro",
                    "route_source": "custom",
                    "item_count": "4",
                    "missing_count": "1",
                    "items": [{"large": "payload"}],
                },
                {"key": ""},
                {"key": "custom--electro", "label": "Duplicate"},
            ],
        }
    )

    assert info["group_overrides"] == {"sparepart:1|PN1": "Electro"}
    assert info["groups"] == [
        {
            "key": "custom--electro",
            "label": "Electro",
            "display_label": "Spare Parts - Electro",
            "document_label": "Spare Parts - Electro",
            "route_source": "custom",
            "item_count": 4,
            "missing_count": 1,
        }
    ]


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


def test_merge_order_state_sections_keeps_unselected_current_values():
    current = {
        "selections": {"production::Cutting": "Current Supplier"},
        "doc_types": {"production::Cutting": "Bestelbon"},
        "doc_numbers": {"production::Cutting": "BB-1"},
        "vat_rates": {"production::Cutting": "21"},
        "pricing": {"production::Cutting": {"unit_price": "1"}},
    }
    incoming = {
        "selections": {"production::Cutting": "Log Supplier"},
        "doc_types": {"production::Cutting": "Offerteaanvraag"},
        "doc_numbers": {"production::Cutting": "OFF-2"},
        "vat_rates": {"production::Cutting": "6"},
        "pricing": {"production::Cutting": {"unit_price": "5"}},
    }

    merged = merge_order_state_sections(current, incoming, {"documents", "pricing", "vat"})

    assert state_keys_for_import_sections({"documents", "pricing", "vat"}) == {
        "doc_types",
        "doc_numbers",
        "pricing",
        "vat_rates",
    }
    assert merged["selections"]["production::Cutting"] == "Current Supplier"
    assert merged["doc_types"]["production::Cutting"] == "Offerteaanvraag"
    assert merged["doc_numbers"]["production::Cutting"] == "OFF-2"
    assert merged["vat_rates"]["production::Cutting"] == "6"
    assert merged["pricing"]["production::Cutting"]["unit_price"] == "5"


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


def test_apply_order_pricing_uses_line_prices_before_fallback():
    items, layout = _apply_order_pricing(
        [
            {"PartNumber": "A", "Description": "Plaat", "Aantal": 2},
            {"PartNumber": "B", "Description": "Buis", "Aantal": 3},
        ],
        {
            "unit_price": "10",
            "items": {
                "part|A": {"unit_price": "12,50", "total_price": ""},
                "part|B": {"unit_price": "", "total_price": "99"},
            },
        },
        context_kind="Productie",
    )

    assert layout is not None
    assert items[0]["Eenheidsprijs"] == "12.50"
    assert items[0]["Totaalprijs"] == "25.00"
    assert items[1]["Eenheidsprijs"] == "10"
    assert items[1]["Totaalprijs"] == "99"


def test_apply_order_pricing_adds_vat_summary_rows():
    items, layout = _apply_order_pricing(
        [{"PartNumber": "A", "Description": "Plaat", "Aantal": 2}],
        {"unit_price": "10"},
        context_kind="Productie",
        vat_rate="21",
    )

    assert layout is not None
    assert items[-3]["Description"] == "Subtotaal excl. BTW"
    assert items[-3]["Totaalprijs"] == "20.00"
    assert items[-2]["Description"] == "BTW 21%"
    assert items[-2]["Totaalprijs"] == "4.20"
    assert items[-1]["Description"] == "Totaal incl. BTW"
    assert items[-1]["Totaalprijs"] == "24.20"


def test_apply_order_pricing_accepts_european_price_formats():
    items, layout = _apply_order_pricing(
        [{"PartNumber": "A", "Description": "Plaat", "Aantal": 2}],
        {"unit_price": "€ 1.234,50"},
        context_kind="Productie",
        vat_rate="21%",
    )

    assert layout is not None
    assert items[0]["Eenheidsprijs"] == "1234.50"
    assert items[0]["Totaalprijs"] == "2469.00"
    assert items[-1]["Totaalprijs"] == "2987.49"


def test_export_log_compatibility_detects_bom_and_selection_differences():
    original_bom = pd.DataFrame(
        [{"PartNumber": "A", "Production": "Cutting", "Aantal": 2}]
    )
    current_bom = pd.DataFrame(
        [{"PartNumber": "B", "Production": "Roof", "Aantal": 1}]
    )
    payload = build_export_session_log(
        project_number="20250165",
        project_name="Piva",
        client_name="Tecno Art bvba",
        bom_source_path="C:/tmp/bom.xlsx",
        bom_df=original_bom,
        state={"selections": {"production::Cutting": "MCB"}},
    )

    summary = summarize_export_log_compatibility(
        payload,
        {"production::Roof"},
        current_bom_df=current_bom,
    )
    message = format_export_log_compatibility_message(summary)

    assert summary["bom_changed"] is True
    assert summary["missing_keys"] == ["production::Cutting"]
    assert summary["new_keys"] == ["production::Roof"]
    assert "Productie: Cutting" in message
    assert "Productie: Roof" in message


def test_export_log_compatibility_formats_spare_part_keys():
    payload = build_export_session_log(
        project_number="20250165",
        project_name="Piva",
        client_name="Tecno Art bvba",
        bom_source_path="C:/tmp/bom.xlsx",
        bom_df=pd.DataFrame([{"PartNumber": "A", "Production": "Spare Parts"}]),
        state={"selections": {"sparepart::supplier--herbaroof": "Herbaroof"}},
    )

    summary = summarize_export_log_compatibility(payload, set())
    message = format_export_log_compatibility_message(summary)

    assert "Spare parts: supplier--herbaroof" in message


def test_export_log_compatibility_marks_logged_spare_part_groups_restorable():
    payload = build_export_session_log(
        project_number="20250165",
        project_name="Piva",
        client_name="Tecno Art bvba",
        bom_source_path="C:/tmp/bom.xlsx",
        bom_df=pd.DataFrame([{"PartNumber": "PN1", "Production": "Spare Parts"}]),
        state={"selections": {"sparepart::custom--electro": "ElectroShop"}},
        spare_parts={
            "group_overrides": {"sparepart:0|PN1": "Electro"},
            "groups": [
                {
                    "key": "custom--electro",
                    "label": "Electro",
                    "display_label": "Spare Parts - Electro",
                    "document_label": "Spare Parts - Electro",
                    "route_source": "custom",
                    "item_count": 1,
                }
            ],
        },
    )

    summary = summarize_export_log_compatibility(payload, set())
    message = format_export_log_compatibility_message(summary)

    assert summary["missing_keys"] == []
    assert summary["restorable_spare_part_keys"] == ["sparepart::custom--electro"]
    assert "via de verdeling hersteld kunnen worden" in message
    assert "Spare parts: custom--electro" in message


def test_resolve_export_document_path_stays_inside_export_dir(tmp_path):
    log_path = tmp_path / "bundle" / EXPORT_SESSION_LOG_FILENAME
    log_path.parent.mkdir()
    inside = resolve_export_document_path(log_path, {"path": "Cutting/order.pdf"})
    outside = resolve_export_document_path(log_path, {"path": "../outside.pdf"})

    assert inside == str(log_path.parent / "Cutting" / "order.pdf")
    assert outside == ""
