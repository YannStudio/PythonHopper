import datetime
import zipfile
from pathlib import Path

import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

from orders import combine_pdfs_per_production, combine_pdfs_from_source


def _blank_pdf(path: Path, *, width: float = 72, height: float = 72) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=width, height=height)
    with path.open("wb") as fh:
        writer.write(fh)


def test_combine_handles_loose_files_and_zip(tmp_path):
    dest = tmp_path
    date = "2023-01-01"
    stamp = datetime.datetime(2023, 1, 1, 12, 30, 15)

    # production with loose PDFs
    prod1 = dest / "prod1"
    prod1.mkdir()
    _blank_pdf(prod1 / "a.pdf")
    _blank_pdf(prod1 / "b.pdf")

    # production with PDFs only inside a zip archive
    prod2 = dest / "prod2"
    prod2.mkdir()
    zip_path = prod2 / "prod2_BB-123.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        for name in ["c.pdf", "d.pdf"]:
            tmp = prod2 / name
            _blank_pdf(tmp)
            zf.write(tmp, arcname=name)
            tmp.unlink()

    # empty production should be skipped
    (dest / "empty").mkdir()

    result = combine_pdfs_per_production(
        str(dest),
        date,
        project_number="PRJ-42",
        project_name="Alpha Beta",
        timestamp=stamp,
    )
    out_dir = Path(result.output_dir)

    assert result.count == 2
    assert out_dir.parent == dest
    folder_name = out_dir.name
    assert folder_name.startswith("Combined pdf_")
    assert "2023-01-01T123015" in folder_name
    assert "PRJ-42" in folder_name
    assert "alpha-beta" in folder_name
    assert sorted(p.name for p in out_dir.glob("*.pdf")) == [
        "prod1_2023-01-01_combined.pdf",
        "prod2_2023-01-01_combined.pdf",
    ]


def test_combine_from_source_without_copy(tmp_path):
    source = tmp_path / "src"
    dest = tmp_path / "out"
    source.mkdir()
    dest.mkdir()

    for name in ["a.pdf", "b.pdf", "c.pdf", "d.pdf"]:
        _blank_pdf(source / name)

    bom_df = pd.DataFrame(
        [
            {"PartNumber": "a", "Production": "prod1"},
            {"PartNumber": "b", "Production": "prod1"},
            {"PartNumber": "c", "Production": "prod2"},
            {"PartNumber": "d", "Production": "prod2"},
        ]
    )

    stamp = datetime.datetime(2023, 1, 1, 13, 45, 0)
    result = combine_pdfs_from_source(
        str(source),
        bom_df,
        str(dest),
        "2023-01-01",
        project_number="PN-5",
        project_name="Gamma",
        timestamp=stamp,
    )
    out_dir = Path(result.output_dir)

    assert result.count == 2
    assert out_dir.parent == dest
    assert "2023-01-01T134500" in out_dir.name
    assert "PN-5" in out_dir.name
    assert "gamma" in out_dir.name
    assert sorted(p.name for p in out_dir.glob("*.pdf")) == [
        "prod1_2023-01-01_combined.pdf",
        "prod2_2023-01-01_combined.pdf",
    ]


def test_combine_from_source_to_dest_without_copy(tmp_path):
    source = tmp_path / "src"
    dest = tmp_path / "out"
    source.mkdir()
    dest.mkdir()

    for name in ["a.pdf", "b.pdf", "c.pdf", "d.pdf"]:
        _blank_pdf(source / name)

    bom_df = pd.DataFrame(
        [
            {"PartNumber": "a", "Production": "prod1"},
            {"PartNumber": "b", "Production": "prod1"},
            {"PartNumber": "c", "Production": "prod2"},
            {"PartNumber": "d", "Production": "prod2"},
        ]
    )

    stamp = datetime.datetime(2023, 1, 1, 9, 0, 0)
    result = combine_pdfs_from_source(
        str(source),
        bom_df,
        str(dest),
        "2023-01-01",
        project_number="PN-5",
        project_name="Gamma",
        timestamp=stamp,
    )
    out_dir = Path(result.output_dir)

    assert result.count == 2
    assert sorted(p.name for p in out_dir.glob("*.pdf")) == [
        "prod1_2023-01-01_combined.pdf",
        "prod2_2023-01-01_combined.pdf",
    ]
    assert [p for p in dest.iterdir()] == [out_dir]


def test_combine_from_source_single_pdf(tmp_path):
    source = tmp_path / "src"
    dest = tmp_path / "out"
    source.mkdir()
    dest.mkdir()

    for name in ["a.pdf", "b.pdf", "c.pdf"]:
        _blank_pdf(source / name)

    bom_df = pd.DataFrame(
        [
            {"PartNumber": "a", "Production": "prod1"},
            {"PartNumber": "b", "Production": "prod1"},
            {"PartNumber": "c", "Production": "prod2"},
        ]
    )

    result = combine_pdfs_from_source(
        str(source),
        bom_df,
        str(dest),
        "2023-01-01",
        combine_per_production=False,
    )

    out_dir = Path(result.output_dir)

    assert result.count == 1
    assert sorted(p.name for p in out_dir.glob("*.pdf")) == [
        "BOM_2023-01-01_combined.pdf",
    ]


def test_combine_from_source_includes_related_bom_pdf(tmp_path):
    source = tmp_path / "src"
    dest = tmp_path / "out"
    bom_source = tmp_path / "Project123-BOM.xlsx"
    source.mkdir()
    dest.mkdir()
    bom_source.write_text("dummy")

    _blank_pdf(source / "Project123.pdf", width=150)
    _blank_pdf(source / "part1.pdf", width=80)
    _blank_pdf(source / "part2.pdf", width=90)

    bom_df = pd.DataFrame(
        [
            {"PartNumber": "part1", "Production": "prod1"},
            {"PartNumber": "part2", "Production": "prod1"},
        ]
    )

    result = combine_pdfs_from_source(
        str(source),
        bom_df,
        str(dest),
        "2023-01-01",
        bom_source_path=str(bom_source),
    )

    out_dir = Path(result.output_dir)
    combined_path = out_dir / "prod1_2023-01-01_combined.pdf"
    reader = PdfReader(str(combined_path))

    # When combining per production, the related BOM PDF should be copied
    # separately to the output directory, not included in the combined PDF
    assert result.count == 1
    assert len(reader.pages) == 2
    assert {float(reader.pages[0].mediabox.width), float(reader.pages[1].mediabox.width)} == {
        80,
        90,
    }
    
    # The related PDF should be copied to the output directory
    related_pdf = out_dir / "Project123.pdf"
    assert related_pdf.exists()


def test_combine_all_pdfs_places_related_first(tmp_path):
    source = tmp_path / "src"
    dest = tmp_path / "out"
    bom_source = tmp_path / "Project123-BOM.xlsx"
    source.mkdir()
    dest.mkdir()
    bom_source.write_text("dummy")

    _blank_pdf(source / "Project123.pdf", width=130)
    _blank_pdf(source / "part1.pdf", width=85)
    _blank_pdf(source / "part2.pdf", width=95)

    bom_df = pd.DataFrame(
        [
            {"PartNumber": "part1", "Production": "prod1"},
            {"PartNumber": "part2", "Production": "prod2"},
        ]
    )

    result = combine_pdfs_from_source(
        str(source),
        bom_df,
        str(dest),
        "2023-01-01",
        combine_per_production=False,
        bom_source_path=str(bom_source),
    )

    out_dir = Path(result.output_dir)
    combined_path = out_dir / "BOM_2023-01-01_combined.pdf"
    reader = PdfReader(str(combined_path))

    assert result.count == 1
    assert len(reader.pages) == 3
    assert float(reader.pages[0].mediabox.width) == 130
    assert {
        float(reader.pages[1].mediabox.width),
        float(reader.pages[2].mediabox.width),
    } == {85, 95}

