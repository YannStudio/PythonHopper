from models import Supplier
from gui import sort_supplier_options


def test_sort_supplier_options_favorites_first():
    sups = [
        Supplier.from_any({"supplier": "Fav1", "favorite": True}),
        Supplier.from_any({"supplier": "Norm", "favorite": False}),
        Supplier.from_any({"supplier": "Fav2", "favorite": True}),
    ]
    disp_to_name = {"★ Fav1": "Fav1", "Norm": "Norm", "★ Fav2": "Fav2"}
    options = ["Norm", "★ Fav2", "★ Fav1"]
    sorted_opts = sort_supplier_options(options, sups, disp_to_name)
    assert sorted_opts == ["★ Fav1", "★ Fav2", "Norm"]


def test_sort_supplier_options_uses_db_not_display_prefix():
    sups = [
        Supplier.from_any({"supplier": "Fav", "favorite": True}),
        Supplier.from_any({"supplier": "Other", "favorite": False}),
    ]
    disp_to_name = {"Fav": "Fav", "Other": "Other"}
    options = ["Other", "Fav"]
    sorted_opts = sort_supplier_options(options, sups, disp_to_name)
    assert sorted_opts == ["Fav", "Other"]
