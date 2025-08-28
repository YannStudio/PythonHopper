from dataclasses import dataclass
from typing import Optional

from helpers import _to_str


@dataclass
class Supplier:
    supplier: str
    description: Optional[str] = None
    supplier_id: Optional[str] = None
    adres_1: Optional[str] = None
    adres_2: Optional[str] = None
    postcode: Optional[str] = None
    gemeente: Optional[str] = None
    land: Optional[str] = None
    btw: Optional[str] = None
    contact_sales: Optional[str] = None
    sales_email: Optional[str] = None
    phone: Optional[str] = None
    favorite: bool = False

    @staticmethod
    def from_any(d: dict) -> "Supplier":
        """Normalize kolomnamen uit CSV/JSON naar canonical fields."""
        key_map = {
            # naam & beschrijving
            "supplier": "supplier",
            "leverancier": "supplier",
            "supplier name": "supplier",
            "naam": "supplier",
            "description": "description",
            "beschrijving": "description",
            "omschrijving": "description",
            "notes": "description",
            # id
            "supplier_id": "supplier_id",
            "supplier id": "supplier_id",
            "id": "supplier_id",
            # adres 1
            "adres_1": "adres_1",
            "adres 1": "adres_1",
            "adres1": "adres_1",
            "address_1": "adres_1",
            "address 1": "adres_1",
            "address": "adres_1",
            "adress_1": "adres_1",
            "adress 1": "adres_1",
            "adress1": "adres_1",
            "straat": "adres_1",
            # adres 2
            "adres_2": "adres_2",
            "adres 2": "adres_2",
            "adres2": "adres_2",
            "address_2": "adres_2",
            "address 2": "adres_2",
            "adress_2": "adres_2",
            "adress 2": "adres_2",
            "adress2": "adres_2",
            # postcode
            "postcode": "postcode",
            "postal code": "postcode",
            "zip": "postcode",
            "zip code": "postcode",
            # gemeente / stad / city / plaats
            "gemeente": "gemeente",
            "stad": "gemeente",
            "city": "gemeente",
            "plaats": "gemeente",
            "town": "gemeente",
            # land
            "land": "land",
            "country": "land",
            # btw (veel varianten)
            "btw": "btw",
            "vat": "btw",
            "btw nummer": "btw",
            "btw-nummer": "btw",
            "btw number": "btw",
            "vat number": "btw",
            "btw number:": "btw",
            "btw nr": "btw",
            "btw nr.": "btw",
            "btw-nr": "btw",
            "btw-nr.": "btw",
            "btw no": "btw",
            "btw no.": "btw",
            "vat no": "btw",
            "vat no.": "btw",
            "vat id": "btw",
            "vat identification number": "btw",
            "vat reg": "btw",
            "vat reg.": "btw",
            "vat reg number": "btw",
            # contact
            "contact sales": "contact_sales",
            "contact_sales": "contact_sales",
            "sales contact": "contact_sales",
            "sales_contact": "contact_sales",
            # email
            "sales e-mail": "sales_email",
            "sales email": "sales_email",
            "e-mail sales": "sales_email",
            "email sales": "sales_email",
            "sales_email": "sales_email",
            "email": "sales_email",
            "mail": "sales_email",
            # phone
            "phone": "phone",
            "phone number": "phone",
            "telefoon": "phone",
            "telefoon nummer": "phone",
            "tel": "phone",
            "tel. sales": "phone",
            "tel sales": "phone",
            # favorite
            "favorite": "favorite",
            "favoriet": "favorite",
            "fav": "favorite",
        }
        norm = {}
        for k, v in d.items():
            lk = str(k).strip().lower()
            if lk in key_map:
                norm[key_map[lk]] = v

        name = str(
            norm.get("supplier")
            or d.get("supplier")
            or d.get("Leverancier")
            or ""
        ).strip()
        if not name or name == "-":
            raise ValueError("Supplier name is missing in record.")

        fav = norm.get("favorite", d.get("favorite", False))
        if isinstance(fav, str):
            fav = fav.strip().lower() in ("1", "true", "yes", "y", "ja")

        return Supplier(
            supplier=name,
            description=_to_str(norm.get("description")).strip() or None if ("description" in norm) else None,
            supplier_id=_to_str(norm.get("supplier_id")).strip() or None if ("supplier_id" in norm) else None,
            adres_1=_to_str(norm.get("adres_1")).strip() or None if ("adres_1" in norm) else None,
            adres_2=_to_str(norm.get("adres_2")).strip() or None if ("adres_2" in norm) else None,
            postcode=_to_str(norm.get("postcode")).strip() or None if ("postcode" in norm) else None,
            gemeente=_to_str(norm.get("gemeente")).strip() or None if ("gemeente" in norm) else None,
            land=_to_str(norm.get("land")).strip() or None if ("land" in norm) else None,
            btw=_to_str(norm.get("btw")).strip() or None if ("btw" in norm) else None,
            contact_sales=_to_str(norm.get("contact_sales")).strip() or None if ("contact_sales" in norm) else None,
            sales_email=_to_str(norm.get("sales_email")).strip() or None if ("sales_email" in norm) else None,
            phone=_to_str(norm.get("phone")).strip() or None if ("phone" in norm) else None,
            favorite=bool(fav),
        )

@dataclass
class Client:
    name: str
    address: Optional[str] = None
    vat: Optional[str] = None
    email: Optional[str] = None
    favorite: bool = False

    @staticmethod
    def from_any(d: dict) -> "Client":
        key_map = {
            "name": "name",
            "client": "name",
            "opdrachtgever": "name",
            "address": "address",
            "adres": "address",
            "btw": "vat",
            "vat": "vat",
            "btw nummer": "vat",
            "btw-nummer": "vat",
            "email": "email",
            "e-mail": "email",
            "mail": "email",
            "favorite": "favorite",
            "favoriet": "favorite",
            "fav": "favorite",
        }
        norm = {}
        for k, v in d.items():
            lk = str(k).strip().lower()
            if lk in key_map:
                norm[key_map[lk]] = v
        name = str(norm.get("name") or d.get("name") or "").strip()
        if not name:
            raise ValueError("Client name is missing in record.")
        fav = norm.get("favorite", d.get("favorite", False))
        if isinstance(fav, str):
            fav = fav.strip().lower() in ("1", "true", "yes", "y", "ja")
        return Client(
            name=name,
            address=_to_str(norm.get("address")).strip() or None if ("address" in norm) else None,
            vat=_to_str(norm.get("vat")).strip() or None if ("vat" in norm) else None,
            email=_to_str(norm.get("email")).strip() or None if ("email" in norm) else None,
            favorite=bool(fav),
        )


@dataclass
class DeliveryAddress:
    """Contactinformatie voor levering."""

    name: str
    address: Optional[str] = None
    contact: Optional[str] = None
    phone: Optional[str] = None
    email: Optional[str] = None
    favorite: bool = False

    @staticmethod
    def from_any(d: dict) -> "DeliveryAddress":
        key_map = {
            "name": "name",
            "naam": "name",
            "address": "address",
            "adres": "address",
            "contact": "contact",
            "contactpersoon": "contact",
            "contact person": "contact",
            "contact name": "contact",
            "phone": "phone",
            "telefoon": "phone",
            "tel": "phone",
            "email": "email",
            "e-mail": "email",
            "mail": "email",
            "favorite": "favorite",
            "favoriet": "favorite",
            "fav": "favorite",
        }
        norm = {}
        for k, v in d.items():
            lk = str(k).strip().lower()
            if lk in key_map:
                norm[key_map[lk]] = v
        name = str(norm.get("name") or d.get("name") or "").strip()
        if not name:
            raise ValueError("Delivery address name is missing in record.")
        fav = norm.get("favorite", d.get("favorite", False))
        if isinstance(fav, str):
            fav = fav.strip().lower() in ("1", "true", "yes", "y", "ja")
        return DeliveryAddress(
            name=name,
            address=_to_str(norm.get("address")).strip() or None if ("address" in norm) else None,
            contact=_to_str(norm.get("contact")).strip() or None if ("contact" in norm) else None,
            phone=_to_str(norm.get("phone")).strip() or None if ("phone" in norm) else None,
            email=_to_str(norm.get("email")).strip() or None if ("email" in norm) else None,
            favorite=bool(fav),
        )
