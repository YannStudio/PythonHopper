from dataclasses import dataclass
from typing import Any, Dict, Optional

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
    product_type: Optional[str] = None
    product_description: Optional[str] = None
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
            # product type en description
            "product / product type": "product_type",
            "product/product type": "product_type",
            "product type": "product_type",
            "producttype": "product_type",
            "product_type": "product_type",
            "category": "product_type",
            "categorie": "product_type",
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
            product_type=_to_str(norm.get("product_type")).strip() or None if ("product_type" in norm) else None,
            product_description=_to_str(norm.get("description")).strip() or None if ("description" in norm) else None,
            favorite=bool(fav),
        )

@dataclass
class Client:
    name: str
    address: Optional[str] = None
    vat: Optional[str] = None
    email: Optional[str] = None
    favorite: bool = False
    logo_path: Optional[str] = None
    logo_crop: Optional[Dict[str, int]] = None

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
            "logo": "logo_path",
            "logo_path": "logo_path",
            "logo file": "logo_path",
            "logo_file": "logo_path",
            "logo crop": "logo_crop",
            "logo_crop": "logo_crop",
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

        def _parse_crop(val: Any) -> Optional[Dict[str, int]]:
            if not val:
                return None
            if isinstance(val, dict):
                norm_keys = {str(k).lower(): v for k, v in val.items()}
                keys = {"left", "top", "right", "bottom"}
                if not keys.issubset(norm_keys.keys()):
                    return None
                try:
                    return {
                        "left": int(float(norm_keys.get("left", 0))),
                        "top": int(float(norm_keys.get("top", 0))),
                        "right": int(float(norm_keys.get("right", 0))),
                        "bottom": int(float(norm_keys.get("bottom", 0))),
                    }
                except Exception:
                    return None
            if isinstance(val, (list, tuple)) and len(val) == 4:
                try:
                    l, t, r, b = [int(float(x)) for x in val]
                    return {"left": l, "top": t, "right": r, "bottom": b}
                except Exception:
                    return None
            if isinstance(val, str):
                parts = [p.strip() for p in val.replace(";", ",").split(",") if p.strip()]
                if len(parts) == 4:
                    try:
                        l, t, r, b = [int(float(x)) for x in parts]
                        return {"left": l, "top": t, "right": r, "bottom": b}
                    except Exception:
                        return None
            return None

        crop = _parse_crop(norm.get("logo_crop", d.get("logo_crop")))
        logo_path = _to_str(norm.get("logo_path", d.get("logo_path")))
        logo_path = logo_path.strip() or None if logo_path is not None else None
        return Client(
            name=name,
            address=_to_str(norm.get("address")).strip() or None if ("address" in norm) else None,
            vat=_to_str(norm.get("vat")).strip() or None if ("vat" in norm) else None,
            email=_to_str(norm.get("email")).strip() or None if ("email" in norm) else None,
            favorite=bool(fav),
            logo_path=logo_path,
            logo_crop=crop,
        )


@dataclass
class DeliveryAddress:
    name: str
    address: Optional[str] = None
    remarks: Optional[str] = None
    favorite: bool = False

    @staticmethod
    def from_any(d: dict) -> "DeliveryAddress":
        key_map = {
            "name": "name",
            "naam": "name",
            "address": "address",
            "adres": "address",
            "remarks": "remarks",
            "opmerking": "remarks",
            "opmerkingen": "remarks",
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
            remarks=_to_str(norm.get("remarks")).strip() or None if ("remarks" in norm) else None,
            favorite=bool(fav),
        )
