from pathlib import Path

from bom import read_csv_flex
from models import Supplier


def test_builtin_supplier_template_exists_and_parses():
    template_path = Path("suppliers_template.csv")
    assert template_path.exists()

    df = read_csv_flex(str(template_path))
    assert not df.empty
    assert {"Postcode", "Gemeente", "Land"}.issubset(df.columns)

    parsed = 0
    ml_coating = None
    for rec in df.to_dict(orient="records"):
        try:
            supplier = Supplier.from_any(rec)
            parsed += 1
            if supplier.supplier == "ML Coating":
                ml_coating = supplier
        except Exception:
            continue

    assert parsed > 10
    assert ml_coating is not None
    assert ml_coating.postcode == "2240"
    assert ml_coating.gemeente == "Zandhoven"
    assert ml_coating.land == "Belgie"
    assert ml_coating.contact_sales is None
