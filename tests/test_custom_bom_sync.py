import pandas as pd
import pytest

from bom_sync import MAIN_BOM_COLUMNS, prepare_custom_bom_for_main


def test_prepare_custom_bom_merges_status_and_trims():
    existing = pd.DataFrame(
        [
            {
                "PartNumber": "PN1",
                "Bestanden gevonden": "ja",
                "Status": "ok",
                "Link": "http://example.com/a",
            },
            {
                "PartNumber": " pn1 ",
                "Bestanden gevonden": "laatste",
                "Status": "gewijzigd",
                "Link": "http://example.com/b",
            },
            {
                "PartNumber": "PN2",
                "Bestanden gevonden": "archief",
                "Status": "oude",
                "Link": "http://example.com/c",
            },
        ]
    )

    custom = pd.DataFrame(
        [
            {
                "PartNumber": " PN1 ",
                "Description": "plaat",
                "Production": "Laser",
                "Materiaal": "S235",
                "Aantal": "1001",
            },
            {
                "PartNumber": "pn2",
                "Description": "buis",
                "Production": "Plooien",
                "Materiaal": "Alu",
                "Aantal": "2",
            },
        ]
    )

    result = prepare_custom_bom_for_main(custom, existing)
    assert list(result.columns) == list(MAIN_BOM_COLUMNS)
    assert len(result) == 2
    assert result.loc[0, "PartNumber"] == "PN1"
    assert result.loc[0, "Bestanden gevonden"] == "laatste"
    assert result.loc[0, "Status"] == "gewijzigd"
    assert result.loc[0, "Link"] == "http://example.com/b"
    assert result.loc[0, "Aantal"] == 999  # begrensd
    assert result.loc[1, "PartNumber"] == "pn2"
    assert result.loc[1, "Bestanden gevonden"] == "archief"
    assert result.loc[1, "Aantal"] == 2


def test_prepare_custom_bom_handles_missing_columns():
    custom = pd.DataFrame(
        [
            {"PartNumber": "PN1", "Aantal": 0},
            {"PartNumber": "PN2", "Aantal": -5},
        ]
    )

    result = prepare_custom_bom_for_main(custom)
    assert list(result.columns) == list(MAIN_BOM_COLUMNS)
    assert result.loc[0, "Aantal"] == 1
    assert result.loc[1, "Aantal"] == 1
    for status_col in ("Bestanden gevonden", "Status", "Link"):
        assert (result[status_col] == "").all()


def test_prepare_custom_bom_missing_partnumber_column_raises():
    custom = pd.DataFrame([{"Description": "plaat"}])
    with pytest.raises(ValueError):
        prepare_custom_bom_for_main(custom)


def test_prepare_custom_bom_requires_non_empty_partnumber():
    custom = pd.DataFrame(
        [
            {"PartNumber": "   ", "Description": "plaat"},
            {"PartNumber": None, "Description": "buis"},
        ]
    )
    with pytest.raises(ValueError):
        prepare_custom_bom_for_main(custom)


def test_prepare_custom_bom_empty_input_returns_empty_dataframe():
    empty = pd.DataFrame(columns=["PartNumber", "Aantal"])
    result = prepare_custom_bom_for_main(empty)
    assert list(result.columns) == list(MAIN_BOM_COLUMNS)
    assert result.empty


def test_prepare_custom_bom_preserves_duplicate_custom_rows():
    existing = pd.DataFrame(
        [
            {
                "PartNumber": "DUP1",
                "Bestanden gevonden": "basis",
                "Status": "gecontroleerd",
                "Link": "http://example.com/x",
            },
            {
                "PartNumber": "NEU",
                "Bestanden gevonden": "initieel",
                "Status": "nieuw",
                "Link": "http://example.com/y",
            },
        ]
    )

    custom = pd.DataFrame(
        [
            {
                "PartNumber": "dup1",
                "Description": "plaat links",
                "Production": "Laser",
                "Materiaal": "S235",
                "Aantal": 3,
            },
            {
                "PartNumber": "dup1 ",
                "Description": "plaat rechts",
                "Production": "Laser",
                "Materiaal": "S355",
                "Aantal": 4,
            },
            {
                "PartNumber": "neu",
                "Description": "buis",
                "Production": "Plooien",
                "Materiaal": "INOX",
                "Aantal": "2",
            },
        ]
    )

    result = prepare_custom_bom_for_main(custom, existing)

    assert list(result.columns) == list(MAIN_BOM_COLUMNS)
    assert len(result) == 3
    dup_rows = result[result["PartNumber"].str.upper() == "DUP1"].reset_index(drop=True)
    assert len(dup_rows) == 2
    assert (dup_rows["Bestanden gevonden"] == "basis").all()
    assert (dup_rows["Status"] == "gecontroleerd").all()
    assert (dup_rows["Link"] == "http://example.com/x").all()
    new_row = result[result["PartNumber"].str.upper() == "NEU"].iloc[0]
    assert new_row["Bestanden gevonden"] == "initieel"
    assert new_row["Status"] == "nieuw"
    assert new_row["Link"] == "http://example.com/y"
