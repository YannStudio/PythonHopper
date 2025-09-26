import io
import zipfile

import pytest

pytest.importorskip("PIL")
from PIL import Image

from clients_db import ClientsDB
from models import Client, Supplier
from orders import generate_pdf_order_platypus, REPORTLAB_OK, write_order_excel
import logo_resolver


def test_clients_db_persists_logo_fields(tmp_path):
    client = Client(
        name="ACME",
        logo_path="client_logos/acme.png",
        logo_crop={"left": 5, "top": 10, "right": 105, "bottom": 210},
    )
    db = ClientsDB([client])
    path = tmp_path / "clients.json"
    db.save(path)

    loaded = ClientsDB.load(path)
    assert loaded.clients
    saved = loaded.clients[0]
    assert saved.logo_path == "client_logos/acme.png"
    assert saved.logo_crop == {"left": 5, "top": 10, "right": 105, "bottom": 210}


@pytest.mark.skipif(not REPORTLAB_OK, reason="ReportLab niet beschikbaar")
def test_generate_pdf_order_includes_logo(tmp_path):
    logo_path = tmp_path / "logo.png"
    Image.new("RGB", (400, 200), "red").save(logo_path)

    pdf_path = tmp_path / "order.pdf"
    company = {
        "name": "ACME",
        "address": "Example Street 1",
        "vat": "BE0123456789",
        "email": "info@example.com",
        "logo_path": str(logo_path),
        "logo_crop": {"left": 50, "top": 0, "right": 350, "bottom": 200},
    }
    supplier = Supplier(supplier="Supplier BV")
    items = [
        {
            "PartNumber": "PN-1",
            "Description": "Onderdeel",
            "Materiaal": "",
            "Aantal": 1,
            "Oppervlakte": "",
            "Gewicht": "",
        }
    ]

    generate_pdf_order_platypus(
        str(pdf_path),
        company,
        supplier,
        production="PROD-1",
        items=items,
    )

    from PyPDF2 import PdfReader

    reader = PdfReader(str(pdf_path))
    page = reader.pages[0]
    resources = page.get("/Resources")
    assert resources is not None
    xobjects = resources.get("/XObject")
    assert xobjects is not None
    dims = []
    for obj in xobjects.values():
        xobj = obj.get_object()
        if xobj.get("/Subtype") == "/Image":
            dims.append((int(xobj.get("/Width")), int(xobj.get("/Height"))))
    assert dims
    assert (300, 200) in dims


@pytest.mark.skipif(not REPORTLAB_OK, reason="ReportLab niet beschikbaar")
def test_generate_pdf_order_logo_resolves_after_chdir(tmp_path, monkeypatch):
    logo_dir = tmp_path / "client_logos"
    logo_dir.mkdir()
    logo_path = logo_dir / "temp.png"
    Image.new("RGB", (160, 80), "blue").save(logo_path)

    monkeypatch.setattr(logo_resolver, "CLIENT_LOGO_DIR", logo_dir)

    work_dir = tmp_path / "elsewhere"
    work_dir.mkdir()
    monkeypatch.chdir(work_dir)

    pdf_path = tmp_path / "order_chdir.pdf"
    company = {
        "name": "ACME",
        "address": "Example Street 1",
        "vat": "BE0123456789",
        "email": "info@example.com",
        "logo_path": "client_logos/temp.png",
    }
    supplier = Supplier(supplier="Supplier BV")
    items = [
        {
            "PartNumber": "PN-1",
            "Description": "Onderdeel",
            "Materiaal": "",
            "Aantal": 1,
            "Oppervlakte": "",
            "Gewicht": "",
        }
    ]

    generate_pdf_order_platypus(
        str(pdf_path),
        company,
        supplier,
        production="PROD-2",
        items=items,
    )

    from PyPDF2 import PdfReader

    reader = PdfReader(str(pdf_path))
    page = reader.pages[0]
    resources = page.get("/Resources")
    assert resources is not None
    xobjects = resources.get("/XObject")
    assert xobjects is not None
    widths = [int(obj.get_object().get("/Width")) for obj in xobjects.values()]
    assert any(w >= 150 for w in widths)


def test_write_order_excel_embeds_logo(tmp_path):
    logo_path = tmp_path / "logo.png"
    Image.new("RGB", (120, 60), "green").save(logo_path)

    xlsx_path = tmp_path / "order.xlsx"
    company = {
        "name": "ACME",
        "address": "Example Street 1",
        "vat": "BE0123456789",
        "email": "info@example.com",
        "logo_path": str(logo_path),
    }
    supplier = Supplier(supplier="Supplier BV")
    items = [
        {
            "PartNumber": "PN-1",
            "Description": "Onderdeel",
            "Materiaal": "",
            "Aantal": 1,
            "Oppervlakte": "",
            "Gewicht": "",
        }
    ]

    write_order_excel(str(xlsx_path), items, company_info=company, supplier=supplier)

    with zipfile.ZipFile(xlsx_path) as zf:
        media_entries = [name for name in zf.namelist() if name.startswith("xl/media/")]
        assert media_entries
        image_data = zf.read(media_entries[0])

    with Image.open(io.BytesIO(image_data)) as embedded:
        embedded.load()
        assert embedded.size == (120, 60)
