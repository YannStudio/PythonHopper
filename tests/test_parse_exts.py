import pytest

from cli import parse_exts


def test_parse_exts_normalizes_wildcards():
    assert parse_exts("*.PDF, dxf", "pdf,dxf") == [".dxf", ".pdf"]


def test_parse_exts_invalid_extension():
    with pytest.raises(ValueError) as exc:
        parse_exts("pdf,exe", "pdf,dxf")
    msg = str(exc.value)
    assert "exe" in msg and "pdf" in msg


def test_parse_exts_step_alias():
    assert parse_exts("step", "stp") == [".step", ".stp"]
