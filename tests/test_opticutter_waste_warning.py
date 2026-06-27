import pytest
import pandas as pd

from opticutter import (
    StockScenarioResult,
    analyse_profiles,
    has_excessive_waste,
)


def test_6200_mm_piece_on_12m_stock_is_excessive_waste() -> None:
    df = pd.DataFrame(
        [
            {
                "PartNumber": "P-001",
                "Description": "Lang profiel",
                "Production": "Cutting",
                "Profile": "Tube 100x100",
                "Length profile": 6200,
                "Materiaal": "S235JR",
                "Aantal": 1,
            }
        ]
    )

    analysis = analyse_profiles(df)
    profile = analysis.profiles[0]

    result_12m = profile.scenarios["12000"]
    assert result_12m.waste_pct == pytest.approx(48.333, abs=0.01)
    assert has_excessive_waste(result_12m)
    assert not has_excessive_waste(profile.scenarios["6000"])


def test_excessive_waste_threshold_is_strictly_above_20_percent() -> None:
    assert not has_excessive_waste(
        StockScenarioResult(
            bars=1,
            waste_mm=2000,
            waste_pct=20.0,
            dropped_pieces=0,
            cuts=0,
        )
    )
    assert has_excessive_waste(
        StockScenarioResult(
            bars=1,
            waste_mm=2001,
            waste_pct=20.01,
            dropped_pieces=0,
            cuts=0,
        )
    )
