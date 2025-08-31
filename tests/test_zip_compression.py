import zipfile
import pandas as pd

from orders import copy_per_production_and_orders
from suppliers_db import SuppliersDB


def test_zipfile_compression(tmp_path):
    if zipfile.zlib is None:
        import pytest

        pytest.skip("ZIP_DEFLATED not available")

    src = tmp_path / "src"
    dst = tmp_path / "dst"
    src.mkdir()
    dst.mkdir()

    # create compressible files
    for pn in ["PN1", "PN2"]:
        (src / f"{pn}.txt").write_text("A" * 1000)

    bom_df = pd.DataFrame(
        [{"PartNumber": "PN1", "Production": "P1", "Aantal": 1},
         {"PartNumber": "PN2", "Production": "P1", "Aantal": 1}]
    )

    db = SuppliersDB()
    copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".txt"],
        db,
        {},
        {},
        False,
        zip_parts=True,
    )

    compressed = dst / "P1" / "P1.zip"
    uncompressed = dst / "P1" / "P1_uncompressed.zip"
    with zipfile.ZipFile(uncompressed, "w", compression=zipfile.ZIP_STORED) as z:
        for pn in ["PN1", "PN2"]:
            z.write(src / f"{pn}.txt", arcname=f"{pn}.txt")

    assert compressed.stat().st_size < uncompressed.stat().st_size
