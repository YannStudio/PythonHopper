from pathlib import Path

from bom import read_csv_flex
from models import Supplier


def test_builtin_supplier_template_exists_and_parses():
    template_path = Path("suppliers_template.csv")
    assert template_path.exists()

    df = read_csv_flex(str(template_path))
    assert not df.empty

    parsed = 0
    for rec in df.to_dict(orient="records"):
        try:
            Supplier.from_any(rec)
            parsed += 1
        except Exception:
            continue

    assert parsed > 10
