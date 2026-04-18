import pytest

from orders import (
    OrderDocumentSection,
    _build_order_pdf_section_story,
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
    assert table._cellvalues[0][4] == "m²"
