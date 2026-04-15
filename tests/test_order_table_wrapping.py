from orders import _wrap_words_to_lines


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
