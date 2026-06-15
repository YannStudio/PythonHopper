from pathlib import Path

import pandas as pd

from gui import (
    _generated_documents_include_spare_part_full_list_pdf,
    _missing_pdf_spare_part_preparation_items,
)
from orders import make_spare_part_selection_key


def test_pdf_workdossier_spare_part_list_preparation_is_optional(tmp_path):
    bom_df = pd.DataFrame(
        [
            {"PartNumber": "ASM-1", "Production": "Assembly"},
            {"PartNumber": "BOLT-1", "Production": "Spare Parts"},
        ]
    )

    assert (
        _missing_pdf_spare_part_preparation_items(
            bom_df,
            [],
            tmp_path,
            include_spare_part_list=False,
        )
        == []
    )


def test_pdf_workdossier_reports_missing_spare_part_full_list(tmp_path):
    bom_df = pd.DataFrame(
        [
            {"PartNumber": "BOLT-1", "Production": " spare-parts "},
        ]
    )

    assert _missing_pdf_spare_part_preparation_items(
        bom_df,
        [],
        tmp_path,
        include_spare_part_list=True,
    ) == ["Spare-parts klaarleglijst"]


def test_pdf_workdossier_accepts_existing_spare_part_full_list_pdf(tmp_path):
    pdf_path = tmp_path / "Spare Parts" / "Standaard bon_Spare Parts klaarleglijst.pdf"
    pdf_path.parent.mkdir()
    pdf_path.write_bytes(b"%PDF-1.4\n%%EOF\n")
    records = [
        {
            "path": str(Path("Spare Parts") / pdf_path.name),
            "kind": "order",
            "format": "pdf",
            "selection_key": make_spare_part_selection_key("full"),
        }
    ]
    bom_df = pd.DataFrame(
        [
            {"PartNumber": "BOLT-1", "Production": "Spare Parts"},
        ]
    )

    assert _generated_documents_include_spare_part_full_list_pdf(records, tmp_path)
    assert (
        _missing_pdf_spare_part_preparation_items(
            bom_df,
            records,
            tmp_path,
            include_spare_part_list=True,
        )
        == []
    )
