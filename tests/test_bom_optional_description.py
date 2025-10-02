import pandas as pd

from bom import load_bom


def test_load_bom_without_description(tmp_path):
    bom_path = tmp_path / "bom.csv"
    pd.DataFrame(
        {
            "PartNumber": ["PN1", "PN2"],
            "Production": ["Laser", "Laser"],
            "Aantal": [2, 1],
        }
    ).to_csv(bom_path, index=False)

    df = load_bom(str(bom_path))

    assert list(df.columns) == [
        "PartNumber",
        "Description",
        "Production",
        "Bestanden gevonden",
        "Status",
        "Materiaal",
        "Aantal",
        "Oppervlakte",
        "Gewicht",
    ]
    assert df["Description"].tolist() == ["PN1", "PN2"]

    grouped = df.groupby("PartNumber")["Aantal"].sum().to_dict()
    assert grouped == {"PN1": 2, "PN2": 1}
