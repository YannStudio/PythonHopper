import ast
import datetime
import math
import os
import pathlib
import tkinter.messagebox as messagebox
import types
from typing import Callable, Dict, Iterable, List, Optional, Tuple

import pandas as pd

from app_settings import FileExtensionSetting
from clients_db import ClientsDB
from delivery_addresses_db import DeliveryAddressesDB
from helpers import _to_str, strip_favorite_marker
from manual_order_tab import ManualOrderTab
from models import Client, DeliveryAddress, Supplier
from orders import (
    DEFAULT_FOOTER_NOTE,
    DEFAULT_QUOTE_FOOTER_NOTE,
    _fit_filename_within_path,
    _normalize_doc_number,
    build_document_export_basename,
    format_document_number_for_display,
)
from suppliers_db import SuppliersDB


class DummyVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class DummyManualTab:
    def __init__(self):
        self.last_doc_number = None

    def set_doc_number(self, value):
        self.last_doc_number = value


def _load_app_class(extra_ns=None):
    source = pathlib.Path("gui.py").read_text(encoding="utf-8")
    mod = ast.parse(source)
    start = next(
        node for node in mod.body if isinstance(node, ast.FunctionDef) and node.name == "start_gui"
    )
    sup_cls = next(
        n for n in start.body if isinstance(n, ast.ClassDef) and n.name == "SupplierSelectionFrame"
    )
    app_cls = next(
        n for n in start.body if isinstance(n, ast.ClassDef) and n.name == "App"
    )
    module_ast = ast.Module(body=[sup_cls, app_cls], type_ignores=[])
    code = compile(module_ast, "<gui_extract>", "exec")
    tk_stub = types.SimpleNamespace(
        Frame=type("Frame", (), {}),
        Tk=type("Tk", (), {}),
        Toplevel=type("Toplevel", (), {}),
        BooleanVar=lambda value=None: None,
        StringVar=lambda value=None: None,
        Label=type("Label", (), {}),
        Entry=type("Entry", (), {}),
        Checkbutton=type("Checkbutton", (), {}),
        Button=type("Button", (), {}),
        LabelFrame=type("LabelFrame", (), {}),
    )
    ttk_stub = types.SimpleNamespace(
        Combobox=type("Combobox", (), {}),
        Treeview=type("Treeview", (), {}),
        Scrollbar=type("Scrollbar", (), {}),
        Style=type("Style", (), {}),
    )
    ns = {
        "tk": tk_stub,
        "ttk": ttk_stub,
        "Callable": Callable,
        "Dict": Dict,
        "Iterable": Iterable,
        "List": List,
        "Optional": Optional,
        "Tuple": Tuple,
        "pd": pd,
        "math": math,
        "os": os,
        "Supplier": Supplier,
        "Client": Client,
        "DeliveryAddress": DeliveryAddress,
        "SuppliersDB": SuppliersDB,
        "DeliveryAddressesDB": DeliveryAddressesDB,
        "_to_str": _to_str,
        "strip_favorite_marker": strip_favorite_marker,
        "FileExtensionSetting": FileExtensionSetting,
        "prepare_custom_bom_for_main": lambda df, _current: df,
        "ManualOrderTab": ManualOrderTab,
        "DEFAULT_FOOTER_NOTE": DEFAULT_FOOTER_NOTE,
        "DEFAULT_QUOTE_FOOTER_NOTE": DEFAULT_QUOTE_FOOTER_NOTE,
        "_normalize_doc_number": _normalize_doc_number,
        "_fit_filename_within_path": _fit_filename_within_path,
        "build_document_export_basename": build_document_export_basename,
        "format_document_number_for_display": format_document_number_for_display,
    }
    if extra_ns:
        ns.update(extra_ns)
    exec(code, ns)
    return ns["App"]


def test_manual_order_export_uses_editor_client_supplier_and_delivery(tmp_path, monkeypatch):
    captured = {}

    def fake_write_order_excel(
        path,
        items,
        company_info,
        supplier,
        delivery,
        doc_type,
        doc_number,
        **kwargs,
    ):
        captured["excel"] = {
            "path": path,
            "items": items,
            "company_info": company_info,
            "supplier": supplier,
            "delivery": delivery,
            "doc_type": doc_type,
            "doc_number": doc_number,
            "kwargs": kwargs,
        }
        pathlib.Path(path).write_text("excel", encoding="utf-8")

    def fake_generate_pdf_order_platypus(
        path,
        company_info,
        supplier,
        context_label,
        items,
        **kwargs,
    ):
        captured["pdf"] = {
            "path": path,
            "company_info": company_info,
            "supplier": supplier,
            "context_label": context_label,
            "items": items,
            "kwargs": kwargs,
        }
        pathlib.Path(path).write_text("pdf", encoding="utf-8")

    App = _load_app_class(
        {
            "write_order_excel": fake_write_order_excel,
            "generate_pdf_order_platypus": fake_generate_pdf_order_platypus,
        }
    )

    monkeypatch.setattr(messagebox, "showinfo", lambda *args, **kwargs: None)
    monkeypatch.setattr(messagebox, "showwarning", lambda *args, **kwargs: None)
    monkeypatch.setattr(messagebox, "showerror", lambda *args, **kwargs: None)
    monkeypatch.setattr(os, "startfile", lambda *args, **kwargs: None, raising=False)

    class DummyApp:
        _current_client = App._current_client
        _resolve_delivery_choice = App._resolve_delivery_choice
        _export_manual_order = App._export_manual_order
        _build_document_export_basename = App._build_document_export_basename
        _format_document_display_number = App._format_document_display_number

        def _document_filename_settings_kwargs(self):
            return {
                "profile": "standard",
                "show_doc_type": True,
                "show_doc_number": True,
                "show_context": True,
                "show_date": True,
                "compact_doc_number": False,
                "separator": "underscore",
            }

        def __init__(self):
            self.client_db = ClientsDB(
                [
                    Client(name="Main Client", address="Mainstraat 1"),
                    Client(
                        name="Editor Client",
                        address="Editorstraat 2",
                        vat="BE0123456789",
                        email="editor@example.com",
                        website="https://editor.example.com",
                    ),
                ]
            )
            self.client_var = DummyVar("Main Client")
            self.db = SuppliersDB([Supplier.from_any({"supplier": "ACME"})])
            self.delivery_db = DeliveryAddressesDB([])
            self.manual_order_tab = DummyManualTab()
            self.dest_folder_var = DummyVar(str(tmp_path))
            self.dest_folder = str(tmp_path)
            self.project_number_var = DummyVar("LEI-STN")
            self.project_name_var = DummyVar("Leien-60x80")
            self.footer_note_var = DummyVar("Order footer")
            self.quote_footer_note_var = DummyVar("Quote footer")
            self.document_display_compact_doc_number_var = DummyVar(0)
            self.status_var = DummyVar("")

    app = DummyApp()
    payload = {
        "doc_type": "Bestelbon",
        "doc_number": "BB-",
        "client": "Editor Client",
        "supplier": "ACME",
        "delivery": ManualOrderTab.CLIENT_ADDRESS_PRESET,
        "context_label": "Leien-60x80",
        "context_kind": "document",
        "remark": "Opmerking",
        "items": [{"description": "Plaat", "quantity": 1}],
        "total_weight": None,
        "column_layout": [{"key": "description", "label": "Omschrijving"}],
    }

    app._export_manual_order(payload)

    today = datetime.date.today().strftime("%Y-%m-%d")
    assert app.manual_order_tab.last_doc_number == ""
    assert app.status_var.get() == f"Handmatige bestelbon opgeslagen in {tmp_path}"

    excel = captured["excel"]
    assert pathlib.Path(excel["path"]).name == f"Bestelbon_Leien-60x80_{today}.xlsx"
    assert excel["company_info"]["name"] == "Editor Client"
    assert excel["company_info"]["address"] == "Editorstraat 2"
    assert excel["company_info"]["website"] == "https://editor.example.com"
    assert excel["supplier"] is not None
    assert excel["supplier"].supplier == "ACME"
    assert excel["delivery"] is not None
    assert excel["delivery"].name == "Editor Client"
    assert excel["delivery"].address == "Editorstraat 2"
    assert excel["doc_number"] is None

    pdf = captured["pdf"]
    assert pathlib.Path(pdf["path"]).name == f"Bestelbon_Leien-60x80_{today}.pdf"
    assert pdf["company_info"]["name"] == "Editor Client"
    assert pdf["company_info"]["website"] == "https://editor.example.com"
    assert pdf["kwargs"]["doc_number"] is None
    assert pdf["kwargs"]["footer_note"] == "Order footer"
    assert pdf["kwargs"]["quote_footer_note"] == "Quote footer"
