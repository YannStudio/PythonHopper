"""Diagnostics for Filehopper runtime data and configuration."""

from __future__ import annotations

import datetime as _dt
import json
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Optional

from app_paths import APP_NAME, APP_VERSION, bundle_root, is_frozen, storage_dir
from app_settings import AppSettings, SETTINGS_FILE
from clients_db import CLIENTS_DB_FILE, ClientsDB
from data_storage import BACKUP_DIR_NAME, backup_existing_file
from delivery_addresses_db import DELIVERY_DB_FILE, DeliveryAddressesDB
from helpers import validate_vat
from order_presets_db import ORDER_PRESETS_DB_FILE, OrderPresetsDB
from pdf_workdossier_presets import PDF_WORKDOSSIER_PRESETS_DB_FILE
from spare_part_presets import SPARE_PART_PRESETS_DB_FILE, SparePartPresetsDB
from suppliers_db import SUPPLIERS_DB_FILE, SuppliersDB


@dataclass(frozen=True)
class DataFileDiagnostics:
    label: str
    filename: str
    path: Path
    exists: bool
    status: str
    count_label: str = "-"
    size_label: str = "-"
    modified_label: str = "-"
    writable_label: str = "-"
    backup_count: int = 0
    latest_backup_label: str = "-"
    error: str = ""


@dataclass(frozen=True)
class DiagnosticReport:
    app_name: str
    app_version: str
    runtime_mode: str
    python_version: str
    bundle_path: Path
    storage_path: Path
    current_working_dir: Path
    data_files: list[DataFileDiagnostics] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class _DataFileSpec:
    label: str
    filename: str
    path: Path
    count_key: str


def data_file_specs() -> list[_DataFileSpec]:
    return [
        _DataFileSpec("Leveranciers", "suppliers_db.json", Path(SUPPLIERS_DB_FILE), "suppliers"),
        _DataFileSpec("Klanten", "clients_db.json", Path(CLIENTS_DB_FILE), "clients"),
        _DataFileSpec("Leveradressen", "delivery_addresses_db.json", Path(DELIVERY_DB_FILE), "addresses"),
        _DataFileSpec("Order-presets", "order_presets.json", Path(ORDER_PRESETS_DB_FILE), "rules"),
        _DataFileSpec("Spare-part presets", "spare_part_presets.json", Path(SPARE_PART_PRESETS_DB_FILE), "rules"),
        _DataFileSpec("PDF-werkdossier presets", "pdf_workdossier_presets.json", Path(PDF_WORKDOSSIER_PRESETS_DB_FILE), "presets"),
        _DataFileSpec("Instellingen", "app_settings.json", Path(SETTINGS_FILE), ""),
    ]


def build_diagnostic_report(settings: Optional[AppSettings] = None) -> DiagnosticReport:
    """Collect runtime status, data-file status, and data validation warnings."""

    active_settings = settings or AppSettings.load()
    files = [_inspect_data_file(spec) for spec in data_file_specs()]
    warnings = _validate_data(active_settings)

    return DiagnosticReport(
        app_name=APP_NAME,
        app_version=APP_VERSION,
        runtime_mode="Executable" if is_frozen() else "Python",
        python_version=sys.version.split()[0],
        bundle_path=bundle_root().resolve(),
        storage_path=storage_dir().resolve(),
        current_working_dir=Path.cwd().resolve(),
        data_files=files,
        warnings=warnings,
    )


def create_data_file_backups() -> list[Path]:
    """Create manual backups for all existing Filehopper data files."""

    backups: list[Path] = []
    for spec in data_file_specs():
        backup = backup_existing_file(spec.path.resolve())
        if backup is not None:
            backups.append(backup)
    return backups


def backup_root() -> Path:
    """Return the directory that contains Filehopper data-file backups."""

    first_path = data_file_specs()[0].path.resolve()
    return first_path.parent / BACKUP_DIR_NAME


def format_report_for_clipboard(report: DiagnosticReport) -> str:
    """Return a compact plain-text diagnostic summary."""

    lines = [
        f"{report.app_name} {report.app_version}",
        f"Runtime: {report.runtime_mode}",
        f"Python: {report.python_version}",
        f"Programmamap: {report.bundle_path}",
        f"Datamap: {report.storage_path}",
        f"Werkmap: {report.current_working_dir}",
        "",
        "Databestanden:",
    ]
    for item in report.data_files:
        lines.append(
            f"- {item.filename}: {item.status}, aantal={item.count_label}, "
            f"gewijzigd={item.modified_label}, backups={item.backup_count}, pad={item.path}"
        )
    lines.append("")
    lines.append("Waarschuwingen:")
    if report.warnings:
        lines.extend(f"- {warning}" for warning in report.warnings)
    else:
        lines.append("- Geen waarschuwingen.")
    return "\n".join(lines)


def _inspect_data_file(spec: _DataFileSpec) -> DataFileDiagnostics:
    path = spec.path.resolve()
    exists = path.exists()
    backup_paths = _list_backups(path)
    latest_backup = _format_datetime(_latest_mtime(backup_paths))
    writable = _writable_label(path)

    if not exists:
        return DataFileDiagnostics(
            label=spec.label,
            filename=spec.filename,
            path=path,
            exists=False,
            status="Ontbreekt",
            writable_label=writable,
            backup_count=len(backup_paths),
            latest_backup_label=latest_backup,
        )
    if not path.is_file():
        return DataFileDiagnostics(
            label=spec.label,
            filename=spec.filename,
            path=path,
            exists=True,
            status="Geen bestand",
            writable_label=writable,
            backup_count=len(backup_paths),
            latest_backup_label=latest_backup,
        )

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        status = "Ongeldige JSON"
        count_label = "-"
        error = str(exc)
    else:
        status = "OK"
        count_label = _count_label(raw, spec.count_key)
        error = ""

    stat = path.stat()
    return DataFileDiagnostics(
        label=spec.label,
        filename=spec.filename,
        path=path,
        exists=True,
        status=status,
        count_label=count_label,
        size_label=_format_size(stat.st_size),
        modified_label=_format_datetime(_dt.datetime.fromtimestamp(stat.st_mtime)),
        writable_label=writable,
        backup_count=len(backup_paths),
        latest_backup_label=latest_backup,
        error=error,
    )


def _validate_data(settings: AppSettings) -> list[str]:
    warnings: list[str] = []

    suppliers_db = SuppliersDB.load(SUPPLIERS_DB_FILE)
    clients_db = ClientsDB.load(CLIENTS_DB_FILE)
    delivery_db = DeliveryAddressesDB.load(DELIVERY_DB_FILE)
    presets_db = OrderPresetsDB.load(ORDER_PRESETS_DB_FILE)
    spare_part_presets_db = SparePartPresetsDB.load(SPARE_PART_PRESETS_DB_FILE)

    warnings.extend(_validate_suppliers(suppliers_db))
    warnings.extend(_validate_clients(clients_db))
    warnings.extend(_validate_delivery_addresses(delivery_db))
    warnings.extend(_validate_presets(presets_db, suppliers_db, clients_db))
    warnings.extend(_validate_spare_part_presets(spare_part_presets_db))
    warnings.extend(_validate_export_folders(settings))

    return warnings


def _validate_suppliers(db: SuppliersDB) -> list[str]:
    warnings: list[str] = []
    names: set[str] = set()
    duplicate_names: set[str] = set()
    supplier_ids: dict[str, list[str]] = {}

    for supplier in db.suppliers:
        name = (supplier.supplier or "").strip()
        key = name.casefold()
        if not name:
            warnings.append("Leverancier zonder naam gevonden.")
        elif key in names:
            duplicate_names.add(name)
        else:
            names.add(key)

        supplier_id = (supplier.supplier_id or "").strip()
        if supplier_id:
            supplier_ids.setdefault(supplier_id.casefold(), []).append(name or "(zonder naam)")

        vat = (supplier.btw or "").strip()
        if vat and not _valid_vat_value(vat):
            warnings.append(f"Leverancier '{name or '?'}' heeft mogelijk ongeldig BTW-nummer: {vat}")

    for name in sorted(duplicate_names):
        warnings.append(f"Dubbele leveranciersnaam gevonden: {name}")

    for supplier_id, supplier_names in sorted(supplier_ids.items()):
        if len(supplier_names) > 1:
            shown = ", ".join(supplier_names)
            warnings.append(f"Dubbele supplier ID '{supplier_id}': {shown}")

    return warnings


def _validate_clients(db: ClientsDB) -> list[str]:
    warnings: list[str] = []
    names: set[str] = set()
    duplicates: set[str] = set()

    for client in db.clients:
        name = (client.name or "").strip()
        key = name.casefold()
        if not name:
            warnings.append("Klant zonder naam gevonden.")
        elif key in names:
            duplicates.add(name)
        else:
            names.add(key)

        vat = (client.vat or "").strip()
        if vat and not _valid_vat_value(vat):
            warnings.append(f"Klant '{name or '?'}' heeft mogelijk ongeldig BTW-nummer: {vat}")

    for name in sorted(duplicates):
        warnings.append(f"Dubbele klantnaam gevonden: {name}")
    return warnings


def _validate_delivery_addresses(db: DeliveryAddressesDB) -> list[str]:
    warnings: list[str] = []
    for address in db.addresses:
        name = (address.name or "").strip()
        if not name:
            warnings.append("Leveradres zonder naam gevonden.")
        if name and not (address.address or "").strip():
            warnings.append(f"Leveradres '{name}' heeft geen adres ingevuld.")
    return warnings


def _validate_presets(
    presets_db: OrderPresetsDB,
    suppliers_db: SuppliersDB,
    clients_db: ClientsDB,
) -> list[str]:
    warnings: list[str] = []
    suppliers = {(supplier.supplier or "").strip().casefold() for supplier in suppliers_db.suppliers}
    clients = {(client.name or "").strip().casefold() for client in clients_db.clients}

    for rule in presets_db.rules:
        supplier = (rule.supplier or "").strip()
        if supplier and supplier.casefold() not in suppliers:
            warnings.append(
                f"Order-preset '{rule.name}' verwijst naar onbekende leverancier: {supplier}"
            )
        client = (rule.client or "").strip()
        if client and client.casefold() not in clients:
            warnings.append(f"Order-preset '{rule.name}' verwijst naar onbekende klant: {client}")
    return warnings


def _validate_spare_part_presets(presets_db: SparePartPresetsDB) -> list[str]:
    warnings: list[str] = []
    seen_names: set[str] = set()
    duplicate_names: set[str] = set()
    active_signatures: dict[tuple[str, str, str], str] = {}

    for rule in presets_db.rules:
        name = (rule.name or "").strip()
        name_key = name.casefold()
        if not name:
            warnings.append("Spare-part preset zonder naam gevonden.")
        elif name_key in seen_names:
            duplicate_names.add(name)
        else:
            seen_names.add(name_key)

        if not (rule.pattern or "").strip():
            warnings.append(f"Spare-part preset '{name or '?'}' heeft geen matchwaarde.")
        if not (rule.target_group or "").strip():
            warnings.append(f"Spare-part preset '{name or '?'}' heeft geen doelgroep.")

        if not rule.enabled:
            continue
        signature = (
            (rule.match_field or "").strip().casefold(),
            (rule.match_type or "").strip().casefold(),
            (rule.pattern or "").strip().casefold(),
        )
        if not all(signature):
            continue
        previous = active_signatures.get(signature)
        if previous and previous != name:
            warnings.append(
                "Overlappende actieve spare-part presets gevonden: "
                f"{previous} en {name or '?'}"
            )
        else:
            active_signatures[signature] = name or "?"

    for name in sorted(duplicate_names):
        warnings.append(f"Dubbele spare-part presetnaam gevonden: {name}")
    return warnings


def _validate_export_folders(settings: AppSettings) -> list[str]:
    warnings: list[str] = []
    source = (settings.source_folder or "").strip()
    dest = (settings.dest_folder or "").strip()
    if source and not Path(source).exists():
        warnings.append(f"Bronmap bestaat niet: {source}")
    if dest:
        dest_path = Path(dest)
        if dest_path.exists() and not dest_path.is_dir():
            warnings.append(f"Doelpad is geen map: {dest}")
        elif not dest_path.exists() and not _parent_writable(dest_path):
            warnings.append(f"Doelmap kan niet aangemaakt worden: {dest}")
        elif dest_path.exists() and not os.access(dest_path, os.W_OK):
            warnings.append(f"Geen schrijfrechten in doelmap: {dest}")
    return warnings


def _count_label(raw: Any, key: str) -> str:
    if not key:
        return "aanwezig"
    if isinstance(raw, list):
        return str(len(raw))
    if isinstance(raw, dict):
        value = raw.get(key)
        if isinstance(value, list):
            return str(len(value))
    return "-"


def _valid_vat_value(value: str) -> bool:
    normalized = "".join(ch for ch in value.upper() if ch.isalnum())
    if normalized in {"NAN", "NONE", "NULL"}:
        return True
    return bool(validate_vat(normalized))


def _list_backups(path: Path) -> list[Path]:
    backup_dir = path.parent / BACKUP_DIR_NAME / path.stem
    if not backup_dir.exists():
        return []
    return sorted((item for item in backup_dir.iterdir() if item.is_file()), key=lambda item: item.stat().st_mtime, reverse=True)


def _latest_mtime(paths: Iterable[Path]) -> Optional[_dt.datetime]:
    latest: Optional[_dt.datetime] = None
    for path in paths:
        try:
            value = _dt.datetime.fromtimestamp(path.stat().st_mtime)
        except OSError:
            continue
        if latest is None or value > latest:
            latest = value
    return latest


def _format_datetime(value: Optional[_dt.datetime]) -> str:
    if value is None:
        return "-"
    return value.strftime("%Y-%m-%d %H:%M")


def _format_size(size: int) -> str:
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.1f} MB"


def _writable_label(path: Path) -> str:
    target = path if path.exists() else path.parent
    try:
        return "Ja" if os.access(target, os.W_OK) else "Nee"
    except OSError:
        return "Nee"


def _parent_writable(path: Path) -> bool:
    parent = path.parent
    while parent and not parent.exists():
        if parent == parent.parent:
            break
        parent = parent.parent
    return parent.exists() and os.access(parent, os.W_OK)
