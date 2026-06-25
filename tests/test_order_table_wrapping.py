import pytest
import orders

from orders import (
    OrderDocumentSection,
    _build_order_excel_section_data,
    _build_order_pdf_section_story,
    _order_title_style,
    _order_table_header_min_width,
    generate_pdf_order_platypus,
    _order_palette,
    _wrap_words_to_lines,
)


def _cell_text(cell):
    if hasattr(cell, "getPlainText"):
        return cell.getPlainText()
    return cell


def _row_texts(row):
    return [_cell_text(cell) for cell in row]


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
    assert table._cellvalues[0][5] == "m\u00b2/st"
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
    header = _row_texts(table._cellvalues[0])
    assert header[5] == "m\u00b2/st"
    assert header[6] == "kg/st"
    assert table._cellvalues[2][1].getPlainText() == "Totaal"
    assert table._cellvalues[2][5].getPlainText() == "2.00"
    assert table._cellvalues[2][6].getPlainText() == "50.00"


def test_order_pdf_section_compacts_price_headers():
    pytest.importorskip("reportlab")
    from reportlab.lib.styles import ParagraphStyle

    section = OrderDocumentSection(
        context_label="Document",
        context_kind="document",
        items=[
            {
                "PartNumber": "PN-1",
                "Description": "Onderdeel",
                "Aantal": 4,
                "Eenheidsprijs": "87.50",
                "Totaalprijs": "350.00",
            }
        ],
        column_layout=[
            {"key": "PartNumber", "label": "Artikel nr.", "weight": 1.6},
            {"key": "Description", "label": "Omschrijving", "weight": 2.4},
            {"key": "Aantal", "label": "Aantal", "numeric": True, "justify": "right", "weight": 0.7},
            {"key": "Eenheidsprijs", "label": "Eenheidsprijs (€)", "numeric": True, "justify": "right", "weight": 0.9},
            {"key": "Totaalprijs", "label": "Totaalprijs (€)", "numeric": True, "justify": "right", "weight": 1.0},
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
    header = _row_texts(table._cellvalues[0])
    assert "Eenheidsprijs (€)" not in header
    assert "Totaalprijs (€)" not in header
    assert header[-2:] == ["Prijs/st. (\u20ac)", "Totaal (\u20ac)"]
    assert table._cellvalues[0][-2].style.alignment == 2
    assert table._cellvalues[0][-1].style.alignment == 2


def test_single_section_pdf_keeps_aantal_header_as_one_word(monkeypatch, tmp_path):
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
        production="Glasklemmen",
        items=[
            {
                "part_number": "XXX-GPH-p01",
                "description": "Outer ring",
                "material": "Inox 304L 2B",
                "quantity": 50,
                "weight": 0.10,
                "unit_price": 14.57,
                "total_price": 728.50,
            }
        ],
        label_kind="document",
        column_layout=[
            {"key": "part_number", "label": "Artikel nr.", "weight": 1.6},
            {"key": "description", "label": "Omschrijving", "weight": 2.6},
            {"key": "material", "label": "Materiaal", "weight": 1.6},
            {
                "key": "quantity",
                "label": "Aantal",
                "justify": "right",
                "numeric": True,
                "integer": True,
                "weight": 0.7,
            },
            {
                "key": "weight",
                "label": "kg",
                "justify": "right",
                "numeric": True,
                "weight": 0.8,
            },
            {
                "key": "unit_price",
                "label": "Prijs/st. (\u20ac)",
                "justify": "right",
                "numeric": True,
                "weight": 1.0,
            },
            {
                "key": "total_price",
                "label": "Totaal (\u20ac)",
                "justify": "right",
                "numeric": True,
                "weight": 1.0,
            },
        ],
    )

    tables = [
        flowable
        for flowable in captured["story"]
        if hasattr(flowable, "_cellvalues")
    ]
    order_table = next(
        table
        for table in tables
        if table._cellvalues and "Aantal" in _row_texts(table._cellvalues[0])
    )
    header = _row_texts(order_table._cellvalues[0])
    qty_idx = header.index("Aantal")
    qty_header = order_table._cellvalues[0][qty_idx]

    assert order_table._colWidths[qty_idx] >= _order_table_header_min_width(
        "Aantal",
        9.5,
    )
    assert getattr(qty_header.style, "splitLongWords", None) == 0


def test_order_excel_section_compacts_price_headers():
    section = OrderDocumentSection(
        context_label="Document",
        context_kind="document",
        items=[
            {
                "Description": "Onderdeel",
                "Aantal": 4,
                "Eenheidsprijs": "87.50",
                "Totaalprijs": "350.00",
            }
        ],
        column_layout=[
            {"key": "Description", "label": "Omschrijving"},
            {"key": "Aantal", "label": "Aantal", "numeric": True},
            {"key": "Eenheidsprijs", "label": "Eenheidsprijs (€)", "numeric": True},
            {"key": "Totaalprijs", "label": "Totaalprijs (€)", "numeric": True},
        ],
    )

    df, _left_cols, _wrap_cols = _build_order_excel_section_data(section)

    assert "Eenheidsprijs (€)" not in df.columns
    assert "Totaalprijs (€)" not in df.columns
    assert list(df.columns)[-2:] == ["Prijs/st. (\u20ac)", "Totaal (\u20ac)"]


def test_order_pdf_section_moves_vat_summary_below_table():
    pytest.importorskip("reportlab")
    from reportlab.lib.styles import ParagraphStyle

    section = OrderDocumentSection(
        context_label="Document",
        context_kind="document",
        items=[
            {
                "Aantal": 2,
                "Oppervlakte": "2.00",
                "Gewicht": "15.00",
                "Eenheidsprijs": "5.00",
                "Totaalprijs": "10.00",
            },
            {
                "Aantal": 10,
                "Oppervlakte": "1.00",
                "Gewicht": "8.00",
                "Eenheidsprijs": "0.30",
                "Totaalprijs": "3.00",
            },
            {"Aantal": "Subtotaal excl. BTW", "Totaalprijs": "13.00"},
            {"Aantal": "BTW 21%", "Totaalprijs": "2.73"},
            {"Aantal": "Totaal incl. BTW", "Totaalprijs": "15.73"},
        ],
        total_surface_m2=14.0,
        total_weight_kg=110.0,
        column_layout=[
            {"key": "Aantal", "label": "Aantal", "numeric": True, "justify": "right", "integer": True},
            {"key": "Oppervlakte", "label": "Oppervlakte", "numeric": True, "justify": "right", "total_surface": True},
            {"key": "Gewicht", "label": "Gewicht (kg)", "numeric": True, "justify": "right", "total_weight": True},
            {"key": "Eenheidsprijs", "label": "Prijs/st. (\u20ac)", "numeric": True, "justify": "right"},
            {"key": "Totaalprijs", "label": "Totaal (\u20ac)", "numeric": True, "justify": "right"},
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

    table = story[0]
    assert len(table._cellvalues) == 4
    assert table._cellvalues[3][1].getPlainText() == "Totaal"
    assert table._cellvalues[3][2].getPlainText() == "14.00"
    assert table._cellvalues[3][3].getPlainText() == "110.00"
    assert table._cellvalues[3][5].getPlainText() == "13.00"
    assert "BTW 21%" not in [
        cell.getPlainText()
        for row in table._cellvalues
        for cell in row
        if hasattr(cell, "getPlainText")
    ]

    summary_table = story[-1]
    assert getattr(story[-2], "height", None) == 8
    assert summary_table._cellvalues[0][0].getPlainText() == "BTW 21%"
    assert summary_table._cellvalues[0][1].getPlainText() == "2.73"
    assert summary_table._cellvalues[1][0].getPlainText() == "Totaal incl. BTW"
    assert summary_table._cellvalues[1][1].getPlainText() == "15.73"


def test_pdf_order_title_scales_down_for_long_descriptions():
    pytest.importorskip("reportlab")
    from reportlab.lib.styles import getSampleStyleSheet

    palette = _order_palette({})
    styles = getSampleStyleSheet()

    normal_style = _order_title_style(
        styles,
        palette,
        "Offerteaanvraag document: Bureau",
        usable_w=500,
    )
    long_style = _order_title_style(
        styles,
        palette,
        "Offerteaanvraag document: Bureau inrichting met extra lange omschrijving",
        usable_w=500,
    )

    assert normal_style.fontSize == 20.0
    assert long_style.fontSize < normal_style.fontSize
    assert long_style.leading == long_style.fontSize + 2.0


def test_order_pdf_header_remark_uses_indented_value_column(monkeypatch, tmp_path):
    pytest.importorskip("reportlab")

    captured = {}

    class FakeDoc:
        def __init__(self, *args, **kwargs):
            pass

        def build(self, story):
            captured["story"] = story

    monkeypatch.setattr(orders, "SimpleDocTemplate", FakeDoc)
    remark = "Materiaal op tekening mag vervangen worden naar Inox 304L 2B."

    generate_pdf_order_platypus(
        str(tmp_path / "bestelbon.pdf"),
        {
            "name": "Tecno Art bvba",
            "address": "Kwade Weide 13, 2920 Kalmthout",
            "vat": "BE0460.973.296",
            "email": "jeroen@tecnoart.be",
        },
        None,
        production="Glasklemmen",
        items=[{"PartNumber": "PN-1", "Description": "Glasklem", "Aantal": 1}],
        doc_number="BB2026171",
        project_number="Stock",
        project_name="Glasklemmen",
        order_remark=remark,
    )

    tables = [
        flowable
        for flowable in captured["story"]
        if hasattr(flowable, "_cellvalues")
    ]
    header_table = next(
        table
        for table in tables
        if len(table._cellvalues) == 1
        and len(table._cellvalues[0]) == 2
        and isinstance(table._cellvalues[0][0], list)
    )
    metadata_table = header_table._cellvalues[0][0][0]
    remark_cell = next(
        row[0]
        for row in metadata_table._cellvalues
        if _cell_text(row[0]).startswith("Opmerking:")
    )

    assert _cell_text(remark_cell) == f"Opmerking: {remark}"
    assert remark_cell.style.leftIndent > 0
    assert remark_cell.style.firstLineIndent < 0
    assert abs(remark_cell.style.leftIndent + remark_cell.style.firstLineIndent) < 0.01
    assert sum(metadata_table._colWidths) < header_table._colWidths[0]


def test_order_excel_section_moves_vat_summary_below_table():
    section = OrderDocumentSection(
        context_label="Document",
        context_kind="document",
        items=[
            {"Description": "Onderdeel", "Aantal": 2, "Totaalprijs": "20.00"},
            {"Description": "Subtotaal excl. BTW", "Totaalprijs": "20.00"},
            {"Description": "BTW 21%", "Totaalprijs": "4.20"},
            {"Description": "Totaal incl. BTW", "Totaalprijs": "24.20"},
        ],
        column_layout=[
            {"key": "Description", "label": "Omschrijving"},
            {"key": "Aantal", "label": "Aantal", "numeric": True},
            {"key": "Totaalprijs", "label": "Totaal (\u20ac)", "numeric": True},
        ],
    )

    df, _left_cols, _wrap_cols = _build_order_excel_section_data(section)

    assert list(df["Omschrijving"]) == [
        "Onderdeel",
        "Totaal",
        "",
        "BTW 21%",
        "Totaal incl. BTW",
    ]
    assert list(df["Totaal (\u20ac)"]) == ["20.00", "20.00", "", "4.20", "24.20"]


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
        if table._cellvalues and "Artikel nr." in _row_texts(table._cellvalues[0])
    )

    header = _row_texts(order_table._cellvalues[0])
    assert "Oppervlakte" not in header
    assert "Gewicht (kg)" not in header
    assert header[4] == "m\u00b2/st"
    assert header[5] == "kg/st"
    assert order_table._cellvalues[2][0].getPlainText() == "Totaal"
    assert order_table._cellvalues[2][4].getPlainText() == "0.42"
    assert order_table._cellvalues[2][5].getPlainText() == "3.23"
