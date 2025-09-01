import sys
import types

# Stub external dependencies used during gui import
sys.modules.setdefault(
    "pandas",
    types.SimpleNamespace(
        read_excel=lambda *a, **k: None,
        read_csv=lambda *a, **k: None,
        DataFrame=object,
    ),
)
openpyxl_mod = types.ModuleType("openpyxl")
openpyxl_styles = types.ModuleType("openpyxl.styles")
openpyxl_styles.Alignment = object
openpyxl_mod.styles = openpyxl_styles
sys.modules.setdefault("openpyxl", openpyxl_mod)
sys.modules.setdefault("openpyxl.styles", openpyxl_styles)

from gui import _filter_options


def test_filter_options_mixed_case_and_tokens():
    options = ["Café du Monde", "Cafe Mundo", "Another"]
    # Mixed case filtering
    assert _filter_options("cAfE", options) == ["Café du Monde", "Cafe Mundo"]
    # Tokenised filtering
    assert _filter_options("cafe monde", options) == ["Café du Monde"]
    # Empty input returns all options
    assert _filter_options("", options) == options


def test_filter_options_diacritics():
    options = ["Bäckerei", "Über Supplier", "Cafe"]
    assert _filter_options("backer", options) == ["Bäckerei"]
    assert _filter_options("uber", options) == ["Über Supplier"]
