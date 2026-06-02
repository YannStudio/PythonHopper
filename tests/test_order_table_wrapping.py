import pytest
import orders

from orders import (
    OrderDocumentSection,
    _build_order_pdf_section_story,
    generate_pdf_order_platypus,
    _order_palette,
    _wrap_words_to_lines,
)


def test_wrap_words_to_lines_limits_description_height():
    text = (
        "Hexagon socket button head screw ISO 7380-1 "
        "M12x50 A2 with extra wording to force wrapping"
    )

    lines = _wrap_words_to_lines(
        text,
        width=80,
        font_name="Helvetica",
        font_size=9,
        max_lines=2,
    )

    assert 1 <= len(lines) <= 2
    assert all(line.strip() for line in lines)
    assert lines[-1].endswith("...")


def test_order_pdf_section_uses_square_meter_header():
    pytest.importorskip("reportlab")
    from reportlab.lib.styles import ParagraphStyle

    section = OrderDocumentSection(
        context_label="Laser",
        context_kind="productie",
        items=[
            {
                "PartNumber": "PN-1",
                "Description": "Onderdeel",
                "Materiaal": "S235",
                "Aantal": 1,
                "Oppervlakte": "1.23",
                "Gewicht": "4.56",
            }
        ],
    )
    story = []

    _build_order_pdf_section_story(
        section,
        story=story,
        usable_w=500,
        palette=_order_palette({}),
        section_title_style=ParagraphStyle("section"),
        show_title=False,
    )

    table = story[-1]
    assert table._cellvalues[0][0] == "Nr."
    assert table._cellvalues[0][5] == "m\u00b2"
    assert table._cellvalues[1][0].getPlainText() == "1"


def test_order_pdf_section_numbers_raw_material_rows_and_keeps_total_label():
    pytest.importorskip("reportlab")
    from reportlab.lib.styles import ParagraphStyle

    section = OrderDocumentSection(
        context_label="Brutemateriaal",
        context_kind="brutemateriaal",
        items=[
            {
                "Profiel": "Koker 40x40",
                "Materiaal": "S235",
                "Lengte": "6000",
                "St.": 2,
                "kg": "25.4",
            },
            {
                "Profiel": "L-profiel 30x30",
                "Materiaal": "S235",
                "Lengte": "3000",
                "St.": 1,
                "kg": "8.1",
            },
        ],
        total_weight_kg=33.5,
    )
    story = []

    _build_order_pdf_section_story(
        section,
        story=story,
        usable_w=500,
        palette=_order_palette({}),
        section_title_style=ParagraphStyle("section"),
        show_title=False,
        start_item_number=4,
    )

    table = story[-1]
    assert table._cellvalues[0][0] == "Nr."
    assert table._cellvalues[1][0].getPlainText() == "4"
    assert table._cellvalues[2][0].getPlainText() == "5"
    assert table._cellvalues[3][0].getPlainText() == ""
    assert table._cellvalues[3][1].getPlainText() == "Totaal"


def test_order_pdf_section_compacts_custom_area_and_weight_headers():
    pytest.importorskip("reportlab")
    from reportlab.lib.styles import ParagraphStyle

    section = OrderDocumentSection(
        context_label="Document",
        context_kind="document",
        items=[
            {
                "part_number": "200426-p01",
                "description": "Test piece - ISO 9606-1:20",
                "material": "S235JR",
                "quantity": 20,
                "Oppervlakte": 0.10,
                "Gewicht": 2.50,
            }
        ],
        total_surface_m2=2.0,
        total_weight_kg=50.0,
        column_layout=[
            {"key": "part_number", "label": "Artikel nr.", "justify": "left", "weight": 1.8},
            {"key": "description", "label": "Omschrijving", "justify": "left", "weight": 2.9},
            {"key": "material", "label": "Materiaal", "justify": "left", "weight": 1.8},
            {"key": "quantity", "label": "Aantal", "justify": "right", "numeric": True, "integer": True, "weight": 0.9},
            {"key": "Oppervlakte", "label": "Oppervlakte", "justify": "right", "numeric": True, "weight": 1.1, "total_surface": True},
            {"key": "Gewicht", "label": "Gewicht (kg)", "justify": "right", "numeric": True, "weight": 1.1, "total_weight": True},
        ],
    )
    story = []

    _build_order_pdf_section_story(
        section,
        story=story,
        usable_w=500,
        palette=_order_palette({}),
        section_title_style=ParagraphStyle("section"),
        show_title=False,
    )

    table = story[-1]
    assert table._cellvalues[0][5] == "m\u00b2"
    assert table._cellvalues[0][6] == "kg"
    assert table._cellvalues[2][1].getPlainText() == "Totaal"
    assert table._cellvalues[2][5].getPlainText() == "2.00"
    assert table._cellvalues[2][6].getPlainText() == "50.00"


def test_single_section_pdf_compacts_custom_area_and_weight_headers(monkeypatch, tmp_path):
    pytest.importorskip("reportlab")

    captured = {}

    class FakeDoc:
        def __init__(self, *args, **kwargs):
            pass

        def build(self, story):
            captured["story"] = story

    monkeypatch.setattr(orders, "SimpleDocTemplate", FakeDoc)

    generate_pdf_order_platypus(
        str(tmp_path / "bestelbon.pdf"),
        {},
        None,
        production="Document",
        items=[
            {
                "part_number": "BB001",
                "description": "Ball gutter",
                "material": "ST235JR",
                "quantity": 1,
                "Oppervlakte": 0.42,
                "Gewicht": 3.23,
            }
        ],
        column_layout=[
            {
                "key": "part_number",
                "label": "Artikel nr.",
                "justify": "left",
                "weight": 1.8,
            },
            {
                "key": "description",
                "label": "Omschrijving",
                "justify": "left",
                "weight": 2.9,
            },
            {
                "key": "material",
                "label": "Materiaal",
                "justify": "left",
                "weight": 1.8,
            },
            {
                "key": "quantity",
                "label": "Aantal",
                "justify": "right",
                "numeric": True,
                "integer": True,
                "weight": 0.9,
            },
            {
                "key": "Oppervlakte",
                "label": "Oppervlakte",
                "justify": "right",
                "numeric": True,
                "weight": 1.1,
                "total_surface": True,
            },
            {
                "key": "Gewicht",
                "label": "Gewicht (kg)",
                "justify": "right",
                "numeric": True,
                "weight": 1.1,
                "total_weight": True,
            },
        ],
        total_surface_m2=0.42,
        total_weight_kg=3.23,
    )

    tables = [
        flowable
        for flowable in captured["story"]
        if hasattr(flowable, "_cellvalues")
    ]
    order_table = next(
        table
        for table in tables
        if table._cellvalues and "Artikel nr." in table._cellvalues[0]
    )

    assert "Oppervlakte" not in order_table._cellvalues[0]
    assert "Gewicht (kg)" not in order_table._cellvalues[0]
    assert order_table._cellvalues[0][4] == "m\u00b2"
    assert order_table._cellvalues[0][5] == "kg"
    assert order_table._cellvalues[2][0].getPlainText() == "Totaal"
    assert order_table._cellvalues[2][4].getPlainText() == "0.42"
    assert order_table._cellvalues[2][5].getPlainText() == "3.23"
