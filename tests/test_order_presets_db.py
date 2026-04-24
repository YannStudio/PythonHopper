from order_presets_db import OrderPresetContext, OrderPresetRule, OrderPresetsDB


def test_order_presets_db_evaluates_highest_priority_first_per_field():
    db = OrderPresetsDB(
        [
            OrderPresetRule(
                name="Algemene laserregel",
                priority=10,
                client="",
                selection_kind="production",
                identifiers=["Laser cutting", "Tube laser cutting"],
                doc_type="Bestelbon",
                delivery="Klantadres",
            ),
            OrderPresetRule(
                name="Klant X leverancier",
                priority=50,
                client="Klant X",
                selection_kind="production",
                identifiers=["Laser cutting"],
                supplier="Laserbedrijf Y",
            ),
        ]
    )

    result = db.evaluate(
        OrderPresetContext(
            client="Klant X",
            selection_kind="production",
            identifier="Laser cutting",
        )
    )

    assert result.supplier == "Laserbedrijf Y"
    assert result.doc_type == "Bestelbon"
    assert result.delivery == "Klantadres"
    assert result.applied_rule_names == [
        "Klant X leverancier",
        "Algemene laserregel",
    ]


def test_order_presets_db_skips_manual_only_rules_for_auto_apply():
    db = OrderPresetsDB(
        [
            OrderPresetRule(
                name="Manuele fallback",
                priority=100,
                auto_apply=False,
                client="Klant X",
                selection_kind="production",
                identifiers=["Laser cutting"],
                supplier="Handmatig BV",
            ),
            OrderPresetRule(
                name="Auto leverancier",
                priority=50,
                auto_apply=True,
                client="Klant X",
                selection_kind="production",
                identifiers=["Laser cutting"],
                supplier="Auto BV",
            ),
        ]
    )

    auto_result = db.evaluate(
        {
            "client": "Klant X",
            "selection_kind": "production",
            "identifier": "Laser cutting",
        },
        auto_apply_only=True,
    )
    manual_result = db.evaluate(
        {
            "client": "Klant X",
            "selection_kind": "production",
            "identifier": "Laser cutting",
        }
    )

    assert auto_result.supplier == "Auto BV"
    assert manual_result.supplier == "Handmatig BV"


def test_order_preset_rule_from_any_supports_nested_match_and_apply():
    rule = OrderPresetRule.from_any(
        {
            "name": "Nested",
            "match": {
                "client": "Klant X",
                "selection_kind": "finish",
                "identifiers": ["Poedercoaten", "Galvaniseren"],
            },
            "apply": {
                "supplier": "Afwerker Z",
                "doc_type": "Offerteaanvraag",
                "delivery": "Geen",
                "remark": "Vraag levertermijn",
                "en1090": True,
            },
        }
    )

    assert rule.client == "Klant X"
    assert rule.selection_kind == "finish"
    assert rule.identifiers == ["Poedercoaten", "Galvaniseren"]
    assert rule.supplier == "Afwerker Z"
    assert rule.doc_type == "Offerteaanvraag"
    assert rule.delivery == "Geen"
    assert rule.remark == "Vraag levertermijn"
    assert rule.en1090 is True
