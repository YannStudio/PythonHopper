import csv

import csv

import pandas as pd

from bom import load_bom
from bom_custom_tab import BOMCustomTab


EXPECTED_COLUMNS = [
    "PartNumber",
    "Description",
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
