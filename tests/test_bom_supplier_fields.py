import csv
from types import SimpleNamespace


import pandas as pd

from bom import load_bom
from bom_custom_tab import BOMCustomTab


EXPECTED_COLUMNS = [
    "PartNumber",
    "Description",
    "Profile",
    "Length profile",
    "Plate thickness",
    "Production",
    "Bestanden gevonden",
    "Status",
    "Materiaal",
    "Supplier",
    "Supplier code",
    "Manufacturer",
    "Manufacturer code",
    "Finish",
    "RAL color",
    "Aantal",
    "Oppervlakte",
    "Gewicht",
]


def test_load_bom_preserves_supplier_fields(tmp_path):
    bom_path = tmp_path / "bom.csv"
    pd.DataFrame(
        {
            "PartNumber": ["PN1"],
            "Description": ["Widget"],
            "Profile": ["  L-100  "],
            "Length profile": [" 2500 "],
            "Plate thickness": [" 8 "],
            "Production": ["Laser"],
            "Supplier": ["  Supplier NV  "],
            "Supplier code": [" SUP-001 "],
            "Manufacturer": [" Maker BV "],
            "Manufacturer code": [" M-42 "],
            "Aantal": [3],
        }
    ).to_csv(bom_path, index=False)

    df = load_bom(str(bom_path))

    assert list(df.columns) == EXPECTED_COLUMNS
    assert df.loc[0, "Supplier"] == "Supplier NV"
    assert df.loc[0, "Supplier code"] == "SUP-001"
    assert df.loc[0, "Manufacturer"] == "Maker BV"
    assert df.loc[0, "Manufacturer code"] == "M-42"
    assert df.loc[0, "Profile"] == "L-100"
    assert df.loc[0, "Length profile"] == "2500"
    assert df.loc[0, "Plate thickness"] == "8"


def test_custom_bom_export_includes_supplier_columns(tmp_path):
    tab = object.__new__(BOMCustomTab)
    export_path = tmp_path / "custom.csv"
    rows = [
        [
            "PN1",
            "Panel",
            "5",
            "L",
            "2500",
            "5",
            "Laser",
            "Steel",
            "Anodized",
            "RAL9005",
            "12",
            "3.4",
            " Supplier BV ",
            " SUP-42 ",
            " Maker BV ",
            " M-007 ",
        ]
    ]

    BOMCustomTab._write_csv(tab, export_path, rows)

    with export_path.open(newline="", encoding="utf-8") as fh:
        reader = csv.reader(fh)
        header = next(reader)
        data = next(reader)

    assert list(BOMCustomTab.HEADERS[-4:]) == [
        "Supplier",
        "Supplier code",
        "Manufacturer",
        "Manufacturer code",
    ]
    assert header == list(BOMCustomTab.HEADERS)
    assert data[-4:] == ["Supplier BV", "SUP-42", "Maker BV", "M-007"]


def test_custom_bom_load_from_main_maps_plate_thickness():
    tab = object.__new__(BOMCustomTab)
    tab.DEFAULT_EMPTY_ROWS = BOMCustomTab.DEFAULT_EMPTY_ROWS
    tab.HEADERS = BOMCustomTab.HEADERS
    tab.MAIN_TO_CUSTOM_COLUMN_MAP = BOMCustomTab.MAIN_TO_CUSTOM_COLUMN_MAP
    captured = {}

    def set_dataframe(self, df, **_kwargs):
        captured["df"] = df.copy(deep=True)

    def no_op(self, *_args, **_kwargs):
        return None

    tab._create_empty_dataframe = BOMCustomTab._create_empty_dataframe.__get__(
        tab, BOMCustomTab
    )
    tab._set_dataframe = set_dataframe.__get__(tab, BOMCustomTab)
    tab._store_baseline_state = no_op.__get__(tab, BOMCustomTab)
    tab.clear_history = no_op.__get__(tab, BOMCustomTab)
    tab._update_status = no_op.__get__(tab, BOMCustomTab)

    BOMCustomTab.load_from_main_dataframe(
        tab,
        pd.DataFrame(
            [
                {
                    "PartNumber": "PN1",
                    "Description": "Plaat",
                    "Plate thickness": "10",
                    "Production": "Laser",
                }
            ]
        ),
    )

    result = captured["df"]
    assert result.loc[0, "PartNumber"] == "PN1"
    assert result.loc[0, "Thickness"] == "10"


def test_custom_bom_push_commits_active_editor_before_building_main_dataframe():
    tab = object.__new__(BOMCustomTab)
    tab.HEADERS = BOMCustomTab.HEADERS
    tab.CUSTOM_TO_MAIN_COLUMN_MAP = BOMCustomTab.CUSTOM_TO_MAIN_COLUMN_MAP
    tab.MAIN_COLUMN_ORDER = BOMCustomTab.MAIN_COLUMN_ORDER
    tab.status_messages = []

    data = pd.DataFrame(
        [["PN1", "Plaat", "1", "", "", "", "Laser", "", "", "", "", "", "", "", "", ""]],
        columns=BOMCustomTab.HEADERS,
    )
    tab.table_model = SimpleNamespace(df=data)

    class FakeTable:
        def _commit_active_edit(self):
            tab.table_model.df.loc[0, "Production"] = "Plooien"
            return True

    captured = {}
    tab.table = FakeTable()
    tab.on_push_to_main = lambda df: captured.setdefault("df", df.copy(deep=True))
    tab._update_status = lambda text: tab.status_messages.append(text)

    BOMCustomTab._push_to_main(tab)

    assert captured["df"].loc[0, "Production"] == "Plooien"
