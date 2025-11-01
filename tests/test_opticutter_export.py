import datetime
import pandas as pd

from opticutter import analyse_profiles
from orders import copy_per_production_and_orders
from suppliers_db import SuppliersDB


def test_opticutter_files_written(tmp_path):
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    df = pd.DataFrame(
        [
            {
                "PartNumber": "P-001",
                "Description": "Profiel A",
                "Production": "Prod1",
                "Profile": "U-80",
                "Length profile": 2500,
                "Materiaal": "Staal",
                "Aantal": 2,
            },
            {
                "PartNumber": "P-002",
                "Description": "Profiel B",
                "Production": "Prod1",
                "Profile": "U-80",
                "Length profile": 1500,
                "Materiaal": "Staal",
                "Aantal": 1,
            },
            {
                "PartNumber": "P-003",
                "Description": "Profiel C",
                "Production": "Prod2",
                "Profile": "L-50",
                "Length profile": 3200,
                "Materiaal": "Alu",
                "Aantal": 3,
            },
        ]
    )

    analysis = analyse_profiles(df)
    choices = {profile.key: profile.best_choice for profile in analysis.profiles}

    copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        SuppliersDB([]),
        {},
        {},
        {},
        False,
        opticutter_analysis=analysis,
        opticutter_choices=choices,
        export_bom=False,
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    prod1_dir = dest / "Prod1"
    scenario_path = prod1_dir / f"Opticutter_Prod1_{today}.xlsx"
    order_path = prod1_dir / f"Bestelbon_brutemateriaal_Prod1_{today}.xlsx"

    assert scenario_path.exists(), "Opticutter scenario workbook ontbreekt"
    assert order_path.exists(), "Bestelbon brutemateriaal workbook ontbreekt"

    scenario_df = pd.read_excel(scenario_path, sheet_name="Scenario")
    assert "Keuze" in scenario_df.columns
    assert scenario_df["Keuze"].astype(str).str.len().max() > 0

    order_df = pd.read_excel(order_path)
    assert "Aantal staven" in order_df.columns
    assert order_df["Aantal staven"].notna().any()
