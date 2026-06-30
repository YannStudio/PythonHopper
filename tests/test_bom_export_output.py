import datetime

import pandas as pd

from models import Supplier
from orders import copy_per_production_and_orders, describe_finish_combo
from opticutter import analyse_profiles
from suppliers_db import SuppliersDB


def _make_db() -> SuppliersDB:
    return SuppliersDB(
        [
            Supplier.from_any({"supplier": "ACME"}),
            Supplier.from_any({"supplier": "SteelCo"}),
        ]
    )


def _basic_bom() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "PartNumber": "PN-001",
                "Description": "Behuizing",
                "Production": "Laser",
                "Profile": "L-100",
                "Length profile": 2500,
                "Materiaal": "Staal",
                "Supplier": "Leverancier",
                "Supplier code": "SUP-001",
                "Manufacturer": "Fabrikant",
                "Manufacturer code": "FAB-001",
                "Finish": "Poedercoat",
                "RAL color": "9005",
                "Oppervlakte": "",  # optional field
                "Gewicht": "",  # optional field
                "Aantal": 2,
                "Bestanden gevonden": "pdf",
                "Status": "✅",
            }
        ]
    )


def test_bom_export_written_with_iso_date(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    bom_path = src / "ProjectX-BOM.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    df = _basic_bom()

    copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
        export_bom=True,
        bom_source_path=str(bom_path),
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    export_path = dest / f"ProjectX-BOM-{today}.xlsx"
    assert export_path.is_file()

    exported = pd.read_excel(export_path)
    expected = df.drop(columns=["Bestanden gevonden", "Status", "Link"], errors="ignore")
    if "Aantal" in expected.columns and "QTY." not in expected.columns:
        expected = expected.rename(columns={"Aantal": "QTY."})
    expected_columns = [
        "PartNumber",
        "Description",
        "QTY.",
        "Profile",
        "Length profile",
        "Production",
        "Materiaal",
        "Supplier",
        "Supplier code",
        "Manufacturer",
        "Manufacturer code",
        "Finish",
        "RAL color",
        "Oppervlakte",
        "Gewicht",
    ]
    for column in expected_columns:
        if column not in expected.columns:
            expected[column] = ""
    expected = expected[expected_columns].copy()
    exported = exported[expected_columns].copy()

    text_columns = {
        "PartNumber",
        "Description",
        "Profile",
        "Length profile",
        "Production",
        "Materiaal",
        "Supplier",
        "Supplier code",
        "Manufacturer",
        "Manufacturer code",
        "Finish",
        "RAL color",
        "Oppervlakte",
        "Gewicht",
    }

    def normalize(df: pd.DataFrame) -> pd.DataFrame:
        normalized = {}
        for column in expected_columns:
            series = df[column]
            if column in text_columns:
                normalized[column] = series.map(
                    lambda v: "" if pd.isna(v) else str(v)
                )
            else:
                normalized[column] = series
        return pd.DataFrame(normalized, columns=expected_columns)

    expected_cmp = normalize(expected)
    exported_cmp = normalize(exported)
    pd.testing.assert_frame_equal(
        expected_cmp.reset_index(drop=True), exported_cmp, check_dtype=False
    )
    assert "Bestanden gevonden" not in exported.columns
    assert "Status" not in exported.columns
    assert "Link" not in exported.columns
    for audit_column in (
        "Besteld bij",
        "Bon type",
        "Bon nummer",
        "EN1090 certificaat besteld",
    ):
        assert audit_column in pd.read_excel(export_path).columns


def test_bom_export_adds_order_audit_columns_without_overwriting_supplier(
    tmp_path,
    monkeypatch,
):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    bom_path = src / "ProjectX-BOM.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    df = _basic_bom()
    df.loc[0, "Supplier"] = ""

    copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {"Laser": "Bestelbon"},
        {"Laser": "BB-101"},
        False,
        export_bom=True,
        bom_source_path=str(bom_path),
        en1090_overrides={"Laser": True},
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    exported = pd.read_excel(dest / f"ProjectX-BOM-{today}.xlsx")
    row = exported.iloc[0]

    assert row["Supplier"] == "" or pd.isna(row["Supplier"])
    assert row["Besteld bij"] == "ACME"
    assert row["Bon type"] == "Bestelbon"
    assert row["Bon nummer"] == "BB-101"
    assert row["EN1090 certificaat besteld"] == "ja"


def test_bom_export_tracks_opticutter_raw_material_order_separately(
    tmp_path,
    monkeypatch,
):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    bom_path = src / "ProjectX-BOM.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    df = pd.DataFrame(
        [
            {
                "PartNumber": "PN-001",
                "Description": "Profiel",
                "Production": "Tube laser",
                "Profile": "U-80",
                "Length profile": 2500,
                "Materiaal": "Staal",
                "Aantal": 2,
                "Gewicht": 5.0,
                "Supplier": "BOM supplier",
            }
        ]
    )
    analysis = analyse_profiles(df)
    choices = {profile.key: profile.best_choice for profile in analysis.profiles}

    copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Tube laser": ""},
        {},
        {},
        False,
        export_bom=True,
        bom_source_path=str(bom_path),
        opticutter_analysis=analysis,
        opticutter_choices=choices,
        opticutter_override_map={"Tube laser": "SteelCo"},
        opticutter_doc_num_map={"Tube laser": "BB-RAW"},
        en1090_overrides={"Tube laser": True},
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    exported = pd.read_excel(dest / f"ProjectX-BOM-{today}.xlsx")
    row = exported.iloc[0]

    assert row["Supplier"] == "BOM supplier"
    assert row["Besteld bij"] == "SteelCo"
    assert row["Bon type"] == "Bestelbon"
    assert row["Bon nummer"] == "BB-RAW"
    assert row["EN1090 certificaat besteld"] == "ja"


def test_bom_export_does_not_mark_stock_opticutter_profiles_as_ordered(
    tmp_path,
    monkeypatch,
):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-STOCK.pdf").write_text("dummy", encoding="utf-8")
    (src / "PN-ORDER.pdf").write_text("dummy", encoding="utf-8")
    bom_path = src / "ProjectX-BOM.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    df = pd.DataFrame(
        [
            {
                "PartNumber": "PN-STOCK",
                "Description": "Kort stockstuk",
                "Production": "Tube laser",
                "Profile": "Koker-20",
                "Length profile": 120,
                "Materiaal": "Staal",
                "Aantal": 4,
                "Gewicht": 0.2,
            },
            {
                "PartNumber": "PN-ORDER",
                "Description": "Te bestellen lengte",
                "Production": "Tube laser",
                "Profile": "U-80",
                "Length profile": 2500,
                "Materiaal": "Staal",
                "Aantal": 2,
                "Gewicht": 5.0,
            },
        ]
    )
    analysis = analyse_profiles(df)
    choices = {
        profile.key: ("stock" if profile.profile == "Koker-20" else profile.best_choice)
        for profile in analysis.profiles
    }

    copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Tube laser": ""},
        {},
        {},
        False,
        export_bom=True,
        bom_source_path=str(bom_path),
        opticutter_analysis=analysis,
        opticutter_choices=choices,
        opticutter_override_map={"Tube laser": "SteelCo"},
        opticutter_doc_num_map={"Tube laser": "BB-RAW"},
        en1090_overrides={"Tube laser": True},
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    exported = pd.read_excel(dest / f"ProjectX-BOM-{today}.xlsx")
    stock_row = exported.loc[exported["PartNumber"] == "PN-STOCK"].iloc[0]
    ordered_row = exported.loc[exported["PartNumber"] == "PN-ORDER"].iloc[0]

    assert pd.isna(stock_row["Besteld bij"]) or stock_row["Besteld bij"] == ""
    assert stock_row["EN1090 certificaat besteld"] == "nee"
    assert ordered_row["Besteld bij"] == "SteelCo"
    assert ordered_row["Bon nummer"] == "BB-RAW"
    assert ordered_row["EN1090 certificaat besteld"] == "ja"


def test_bom_export_combines_production_and_finish_order_audit(
    tmp_path,
    monkeypatch,
):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    bom_path = src / "ProjectX-BOM.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    df = _basic_bom()
    finish_key = describe_finish_combo("Poedercoat", "9005")["key"]

    copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {"Laser": "Bestelbon"},
        {"Laser": "BB-101"},
        False,
        export_bom=True,
        bom_source_path=str(bom_path),
        copy_finish_exports=True,
        finish_override_map={finish_key: "SteelCo"},
        finish_doc_num_map={finish_key: "BB-202"},
        en1090_overrides={"Laser": True},
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    exported = pd.read_excel(dest / f"ProjectX-BOM-{today}.xlsx")
    row = exported.iloc[0]

    assert row["Besteld bij"] == "Productie: ACME; Afwerking: SteelCo"
    assert row["Bon type"] == "Bestelbon"
    assert row["Bon nummer"] == "Productie: BB-101; Afwerking: BB-202"
    assert row["EN1090 certificaat besteld"] == "ja"


def test_bom_export_can_be_disabled(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    bom_path = src / "ProjectX-BOM.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    df = _basic_bom()

    copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
        export_bom=False,
        bom_source_path=str(bom_path),
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    export_path = dest / f"ProjectX-BOM-{today}.xlsx"
    assert not export_path.exists()


def test_bom_export_strips_suffix_after_bom(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    bom_path = src / "154215-a21-BOM-partsonly.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    df = _basic_bom()

    copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
        export_bom=True,
        bom_source_path=str(bom_path),
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    export_path = dest / f"154215-a21-BOM-{today}.xlsx"
    assert export_path.exists()


def test_related_exports_copied_next_to_bom(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    assembly_stem = "20230171-v2-a00"
    (src / f"{assembly_stem}.pdf").write_text("top", encoding="utf-8")
    bom_path = src / f"{assembly_stem}-BOM-PartsOnly.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    df = _basic_bom()

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
        export_bom=True,
        bom_source_path=str(bom_path),
    )

    assert cnt == 2
    assert (dest / f"{assembly_stem}.pdf").is_file()


def test_related_exports_copied_without_processed_bom_export(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    assembly_stem = "20230171-v2-a00"
    (src / f"{assembly_stem}.pdf").write_text("top", encoding="utf-8")
    bom_path = src / f"{assembly_stem}-BOM-PartsOnly.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        _basic_bom(),
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
        export_bom=False,
        export_related_files=True,
        bom_source_path=str(bom_path),
    )

    assert cnt == 2
    assert not any(dest.glob("*.xlsx"))
    assert (dest / f"{assembly_stem}.pdf").is_file()


def test_related_exports_match_variant_bom_file_stem(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    (src / "ProjectX-BOM-revA.pdf").write_text("top", encoding="utf-8")
    bom_path = src / "ProjectX-BOM.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    copy_per_production_and_orders(
        str(src),
        str(dest),
        _basic_bom(),
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
        export_bom=True,
        export_related_files=True,
        bom_source_path=str(bom_path),
    )

    assert (dest / "ProjectX-BOM-revA.pdf").is_file()
