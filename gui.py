import os
import shutil
import subprocess
import sys
import threading
import unicodedata
from typing import Dict, List, Optional

import pandas as pd

from PyQt6 import QtCore, QtGui, QtWidgets

from helpers import _build_file_index
from models import Supplier, Client, DeliveryAddress
from suppliers_db import SuppliersDB, SUPPLIERS_DB_FILE
from clients_db import ClientsDB, CLIENTS_DB_FILE
from delivery_addresses_db import DeliveryAddressesDB, DELIVERY_DB_FILE
from bom import read_csv_flex, load_bom
from orders import (
    copy_per_production_and_orders,
    DEFAULT_FOOTER_NOTE,
    combine_pdfs_per_production,
    combine_pdfs_from_source,
    _prefix_for_doc_type,
)


def _norm(text: str) -> str:
    return (
        unicodedata.normalize("NFKD", text)
        .encode("ASCII", "ignore")
        .decode("ASCII")
        .lower()
    )


def sort_supplier_options(
    options: List[str],
    suppliers: List[Supplier],
    disp_to_name: Dict[str, str],
) -> List[str]:
    """Return options sorted with favorites first and then alphabetically.

    Parameters
    ----------
    options: list of display strings
    suppliers: list of Supplier objects from the DB
    disp_to_name: mapping from display string to supplier name
    """

    fav_map = {_norm(s.supplier): s.favorite for s in suppliers}

    def sort_key(opt: str):
        name = disp_to_name.get(opt, opt)
        n = _norm(name)
        return (not fav_map.get(n, False), n)

    return sorted(options, key=sort_key)



class ClientsManagerWidget(QtWidgets.QWidget):
    def __init__(self, db, on_change=None, parent=None):
        super().__init__(parent)
        self.db = db
        self.on_change = on_change

        layout = QtWidgets.QVBoxLayout(self)

        self.table = QtWidgets.QTableWidget(0, 4, self)
        self.table.setHorizontalHeaderLabels(["Naam", "Adres", "BTW", "E-mail"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.table.doubleClicked.connect(self.edit_selected)
        layout.addWidget(self.table)

        btn_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(btn_layout)
        add_btn = QtWidgets.QPushButton("Toevoegen", self)
        add_btn.clicked.connect(self.add_client)
        btn_layout.addWidget(add_btn)
        edit_btn = QtWidgets.QPushButton("Bewerken", self)
        edit_btn.clicked.connect(self.edit_selected)
        btn_layout.addWidget(edit_btn)
        remove_btn = QtWidgets.QPushButton("Verwijderen", self)
        remove_btn.clicked.connect(self.remove_selected)
        btn_layout.addWidget(remove_btn)
        fav_btn = QtWidgets.QPushButton("Favoriet ★", self)
        fav_btn.clicked.connect(self.toggle_favorite)
        btn_layout.addWidget(fav_btn)
        import_btn = QtWidgets.QPushButton("Importeer CSV", self)
        import_btn.clicked.connect(self.import_csv)
        btn_layout.addWidget(import_btn)
        btn_layout.addStretch(1)

        self.refresh()

    def _selected_name(self):
        items = self.table.selectionModel().selectedRows()
        if not items:
            return None
        row = items[0].row()
        name_item = self.table.item(row, 0)
        if not name_item:
            return None
        return name_item.text().replace("★ ", "", 1)

    def refresh(self):
        self.table.setRowCount(0)
        for client in self.db.clients_sorted():
            row = self.table.rowCount()
            self.table.insertRow(row)
            display = self.db.display_name(client)
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(display))
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(client.address or ""))
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(client.vat or ""))
            self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(client.email or ""))
        self.table.resizeColumnsToContents()

    def _open_dialog(self, client=None):
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Opdrachtgever")
        form = QtWidgets.QFormLayout(dialog)
        name_edit = QtWidgets.QLineEdit(dialog)
        address_edit = QtWidgets.QLineEdit(dialog)
        vat_edit = QtWidgets.QLineEdit(dialog)
        email_edit = QtWidgets.QLineEdit(dialog)
        fav_check = QtWidgets.QCheckBox("Favoriet", dialog)
        if client:
            name_edit.setText(client.name or "")
            address_edit.setText(client.address or "")
            vat_edit.setText(client.vat or "")
            email_edit.setText(client.email or "")
            fav_check.setChecked(bool(client.favorite))
        form.addRow("Naam:", name_edit)
        form.addRow("Adres:", address_edit)
        form.addRow("BTW:", vat_edit)
        form.addRow("E-mail:", email_edit)
        form.addRow(fav_check)
        btn_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save | QtWidgets.QDialogButtonBox.StandardButton.Cancel,
            parent=dialog,
        )
        form.addRow(btn_box)

        def on_accept():
            rec = {
                "name": name_edit.text().strip(),
                "address": address_edit.text().strip() or None,
                "vat": vat_edit.text().strip() or None,
                "email": email_edit.text().strip() or None,
                "favorite": fav_check.isChecked(),
            }
            if not rec["name"]:
                QtWidgets.QMessageBox.warning(dialog, "Let op", "Naam is verplicht.")
                return
            from models import Client

            client_obj = Client.from_any(rec)
            self.db.upsert(client_obj)
            from clients_db import CLIENTS_DB_FILE

            self.db.save(CLIENTS_DB_FILE)
            self.refresh()
            if self.on_change:
                self.on_change()
            dialog.accept()

        btn_box.accepted.connect(on_accept)
        btn_box.rejected.connect(dialog.reject)
        dialog.exec()

    def add_client(self):
        self._open_dialog(None)

    def edit_selected(self):
        name = self._selected_name()
        if not name:
            return
        client = self.db.get(name)
        if client:
            self._open_dialog(client)

    def remove_selected(self):
        name = self._selected_name()
        if not name:
            return
        if QtWidgets.QMessageBox.question(
            self,
            "Bevestigen",
            f"Verwijder '{name}'?",
        ) == QtWidgets.QMessageBox.StandardButton.Yes:
            if self.db.remove(name):
                from clients_db import CLIENTS_DB_FILE

                self.db.save(CLIENTS_DB_FILE)
                self.refresh()
                if self.on_change:
                    self.on_change()

    def toggle_favorite(self):
        name = self._selected_name()
        if not name:
            return
        if self.db.toggle_fav(name):
            from clients_db import CLIENTS_DB_FILE

            self.db.save(CLIENTS_DB_FILE)
            self.refresh()
            if self.on_change:
                self.on_change()

    def import_csv(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "CSV of Excel",
            "",
            "CSV (*.csv);;Excel (*.xlsx *.xls)",
        )
        if not path:
            return
        import pandas as pd
        from clients_db import CLIENTS_DB_FILE
        from bom import read_csv_flex
        from models import Client

        try:
            if path.lower().endswith((".xls", ".xlsx")):
                df = pd.read_excel(path)
            else:
                try:
                    df = pd.read_csv(path, encoding="latin1", sep=";")
                except Exception:
                    df = read_csv_flex(path)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Fout", str(exc))
            return
        changed = 0
        for _, row in df.iterrows():
            try:
                rec = {k: row[k] for k in df.columns if k in row}
                client = Client.from_any(rec)
                self.db.upsert(client)
                changed += 1
            except Exception:
                continue
        self.db.save(CLIENTS_DB_FILE)
        self.refresh()
        if self.on_change:
            self.on_change()
        QtWidgets.QMessageBox.information(self, "Import", f"Verwerkt (upsert): {changed}")


class DeliveryAddressesManagerWidget(QtWidgets.QWidget):
    def __init__(self, db, on_change=None, parent=None):
        super().__init__(parent)
        self.db = db
        self.on_change = on_change

        layout = QtWidgets.QVBoxLayout(self)
        self.table = QtWidgets.QTableWidget(0, 3, self)
        self.table.setHorizontalHeaderLabels(["Naam", "Adres", "Opmerkingen"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.table.doubleClicked.connect(self.edit_selected)
        layout.addWidget(self.table)

        btn_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(btn_layout)
        add_btn = QtWidgets.QPushButton("Toevoegen", self)
        add_btn.clicked.connect(self.add_address)
        btn_layout.addWidget(add_btn)
        edit_btn = QtWidgets.QPushButton("Bewerken", self)
        edit_btn.clicked.connect(self.edit_selected)
        btn_layout.addWidget(edit_btn)
        remove_btn = QtWidgets.QPushButton("Verwijderen", self)
        remove_btn.clicked.connect(self.remove_selected)
        btn_layout.addWidget(remove_btn)
        fav_btn = QtWidgets.QPushButton("Favoriet ★", self)
        fav_btn.clicked.connect(self.toggle_favorite)
        btn_layout.addWidget(fav_btn)
        btn_layout.addStretch(1)

        self.refresh()

    def _selected_name(self):
        items = self.table.selectionModel().selectedRows()
        if not items:
            return None
        row = items[0].row()
        name_item = self.table.item(row, 0)
        if not name_item:
            return None
        return name_item.text().replace("★ ", "", 1)

    def refresh(self):
        self.table.setRowCount(0)
        for addr in self.db.addresses_sorted():
            row = self.table.rowCount()
            self.table.insertRow(row)
            display = self.db.display_name(addr)
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(display))
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(addr.address or ""))
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(addr.remarks or ""))
        self.table.resizeColumnsToContents()

    def _open_dialog(self, addr=None):
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Leveradres")
        form = QtWidgets.QFormLayout(dialog)
        name_edit = QtWidgets.QLineEdit(dialog)
        address_edit = QtWidgets.QLineEdit(dialog)
        remarks_edit = QtWidgets.QLineEdit(dialog)
        fav_check = QtWidgets.QCheckBox("Favoriet", dialog)
        if addr:
            name_edit.setText(addr.name or "")
            address_edit.setText(addr.address or "")
            remarks_edit.setText(addr.remarks or "")
            fav_check.setChecked(bool(addr.favorite))
        form.addRow("Naam:", name_edit)
        form.addRow("Adres:", address_edit)
        form.addRow("Opmerkingen:", remarks_edit)
        form.addRow(fav_check)
        btn_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save | QtWidgets.QDialogButtonBox.StandardButton.Cancel,
            parent=dialog,
        )
        form.addRow(btn_box)

        def on_accept():
            rec = {
                "name": name_edit.text().strip(),
                "address": address_edit.text().strip() or None,
                "remarks": remarks_edit.text().strip() or None,
                "favorite": fav_check.isChecked(),
            }
            if not rec["name"]:
                QtWidgets.QMessageBox.warning(dialog, "Let op", "Naam is verplicht.")
                return
            from delivery_addresses_db import DeliveryAddress, DELIVERY_DB_FILE

            addr_obj = DeliveryAddress.from_any(rec)
            self.db.upsert(addr_obj)
            self.db.save(DELIVERY_DB_FILE)
            self.refresh()
            if self.on_change:
                self.on_change()
            dialog.accept()

        btn_box.accepted.connect(on_accept)
        btn_box.rejected.connect(dialog.reject)
        dialog.exec()

    def add_address(self):
        self._open_dialog(None)

    def edit_selected(self):
        name = self._selected_name()
        if not name:
            return
        addr = self.db.get(name)
        if addr:
            self._open_dialog(addr)

    def remove_selected(self):
        name = self._selected_name()
        if not name:
            return
        if QtWidgets.QMessageBox.question(
            self,
            "Bevestigen",
            f"Verwijder '{name}'?",
        ) == QtWidgets.QMessageBox.StandardButton.Yes:
            if self.db.remove(name):
                from delivery_addresses_db import DELIVERY_DB_FILE

                self.db.save(DELIVERY_DB_FILE)
                self.refresh()
                if self.on_change:
                    self.on_change()

    def toggle_favorite(self):
        name = self._selected_name()
        if not name:
            return
        if self.db.toggle_fav(name):
            from delivery_addresses_db import DELIVERY_DB_FILE

            self.db.save(DELIVERY_DB_FILE)
            self.refresh()
            if self.on_change:
                self.on_change()


class SuppliersManagerWidget(QtWidgets.QWidget):
    def __init__(self, db, on_change=None, parent=None):
        super().__init__(parent)
        self.db = db
        self.on_change = on_change

        layout = QtWidgets.QVBoxLayout(self)

        search_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(search_layout)
        search_layout.addWidget(QtWidgets.QLabel("Zoek:", self))
        self.search_edit = QtWidgets.QLineEdit(self)
        self.search_edit.textChanged.connect(self.refresh)
        search_layout.addWidget(self.search_edit, 1)

        self.table = QtWidgets.QTableWidget(0, 6, self)
        self.table.setHorizontalHeaderLabels([
            "Supplier",
            "BTW",
            "E-mail",
            "Tel",
            "Adres 1",
            "Adres 2",
        ])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.table.doubleClicked.connect(self.edit_selected)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)

        btn_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(btn_layout)
        add_btn = QtWidgets.QPushButton("Toevoegen", self)
        add_btn.clicked.connect(self.add_supplier)
        btn_layout.addWidget(add_btn)
        edit_btn = QtWidgets.QPushButton("Bewerken", self)
        edit_btn.clicked.connect(self.edit_selected)
        btn_layout.addWidget(edit_btn)
        remove_btn = QtWidgets.QPushButton("Verwijderen", self)
        remove_btn.clicked.connect(self.remove_selected)
        btn_layout.addWidget(remove_btn)
        merge_btn = QtWidgets.QPushButton("Update uit CSV (merge)", self)
        merge_btn.clicked.connect(self.merge_csv)
        btn_layout.addWidget(merge_btn)
        fav_btn = QtWidgets.QPushButton("Favoriet ★", self)
        fav_btn.clicked.connect(self.toggle_favorite)
        btn_layout.addWidget(fav_btn)
        btn_layout.addStretch(1)

        self.refresh()

    def _selected_name(self):
        items = self.table.selectionModel().selectedRows()
        if not items:
            return None
        row = items[0].row()
        item = self.table.item(row, 0)
        if not item:
            return None
        return item.text().replace("★ ", "", 1)

    def refresh(self):
        query = self.search_edit.text()
        suppliers = self.db.find(query)
        self.table.setRowCount(0)
        for supplier in suppliers:
            row = self.table.rowCount()
            self.table.insertRow(row)
            display = self.db.display_name(supplier)
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(display))
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(supplier.btw or ""))
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(supplier.sales_email or ""))
            self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(supplier.phone or ""))
            self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(supplier.adres_1 or ""))
            self.table.setItem(row, 5, QtWidgets.QTableWidgetItem(supplier.adres_2 or ""))
        self.table.resizeColumnsToContents()

    def add_supplier(self):
        name, ok = QtWidgets.QInputDialog.getText(self, "Nieuwe leverancier", "Naam:")
        if not ok or not name.strip():
            return
        from models import Supplier
        from suppliers_db import SUPPLIERS_DB_FILE

        supplier = Supplier.from_any({"supplier": name.strip()})
        self.db.upsert(supplier)
        self.db.save(SUPPLIERS_DB_FILE)
        self.refresh()
        if self.on_change:
            self.on_change()

    def remove_selected(self):
        name = self._selected_name()
        if not name:
            return
        if QtWidgets.QMessageBox.question(
            self,
            "Bevestigen",
            f"Verwijder '{name}'?",
        ) == QtWidgets.QMessageBox.StandardButton.Yes:
            from suppliers_db import SUPPLIERS_DB_FILE

            if self.db.remove(name):
                self.db.save(SUPPLIERS_DB_FILE)
                self.refresh()
                if self.on_change:
                    self.on_change()

    def toggle_favorite(self):
        name = self._selected_name()
        if not name:
            return
        from suppliers_db import SUPPLIERS_DB_FILE

        if self.db.toggle_fav(name):
            self.db.save(SUPPLIERS_DB_FILE)
            self.refresh()
            if self.on_change:
                self.on_change()

    def merge_csv(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "CSV bestand",
            "",
            "CSV (*.csv);;Alle bestanden (*.*)",
        )
        if not path:
            return
        from bom import read_csv_flex
        from models import Supplier
        from suppliers_db import SUPPLIERS_DB_FILE

        try:
            df = read_csv_flex(path)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Fout", str(exc))
            return
        for rec in df.to_dict(orient="records"):
            try:
                sup = Supplier.from_any(rec)
                self.db.upsert(sup)
            except Exception:
                continue
        self.db.save(SUPPLIERS_DB_FILE)
        self.refresh()
        if self.on_change:
            self.on_change()

    def _edit_dialog(self, supplier):
        from models import Supplier

        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Leverancier bewerken")
        form = QtWidgets.QFormLayout(dialog)
        fields = [
            ("supplier", "Naam"),
            ("description", "Beschrijving"),
            ("supplier_id", "ID"),
            ("adres_1", "Adres 1"),
            ("adres_2", "Adres 2"),
            ("postcode", "Postcode"),
            ("gemeente", "Gemeente"),
            ("land", "Land"),
            ("btw", "BTW"),
            ("contact_sales", "Contact"),
            ("sales_email", "E-mail"),
            ("phone", "Tel"),
        ]
        edits = {}
        for key, label in fields:
            edit = QtWidgets.QLineEdit(dialog)
            edit.setText(getattr(supplier, key) or "")
            form.addRow(label + ":", edit)
            edits[key] = edit
        btn_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save | QtWidgets.QDialogButtonBox.StandardButton.Cancel,
            parent=dialog,
        )
        form.addRow(btn_box)

        def on_accept():
            data = {k: v.text().strip() or None for k, v in edits.items()}
            try:
                updated = Supplier.from_any(data)
            except Exception as exc:
                QtWidgets.QMessageBox.critical(dialog, "Fout", str(exc))
                return
            self.db.upsert(updated)
            from suppliers_db import SUPPLIERS_DB_FILE

            self.db.save(SUPPLIERS_DB_FILE)
            self.refresh()
            if self.on_change:
                self.on_change()
            dialog.accept()

        btn_box.accepted.connect(on_accept)
        btn_box.rejected.connect(dialog.reject)
        dialog.exec()

    def edit_selected(self):
        name = self._selected_name()
        if not name:
            return
        supplier = None
        for s in self.db.suppliers:
            if s.supplier == name:
                supplier = s
                break
        if supplier:
            self._edit_dialog(supplier)


class SupplierSelectionDialog(QtWidgets.QDialog):
    def __init__(self, productions, suppliers_db, delivery_db, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Selecteer leveranciers")
        self.db = suppliers_db
        self.delivery_db = delivery_db
        self.productions = productions

        self.resize(900, 500)

        self.sel_combos = {}
        self.doc_combos = {}
        self.doc_num_edits = {}
        self.delivery_combos = {}

        main_layout = QtWidgets.QVBoxLayout(self)

        project_group = QtWidgets.QGroupBox("Projectgegevens", self)
        project_layout = QtWidgets.QGridLayout(project_group)
        main_layout.addWidget(project_group)
        project_layout.addWidget(QtWidgets.QLabel("Projectnr."), 0, 0)
        self.project_number_edit = QtWidgets.QLineEdit(self)
        project_layout.addWidget(self.project_number_edit, 0, 1)
        project_layout.addWidget(QtWidgets.QLabel("Projectnaam"), 1, 0)
        self.project_name_edit = QtWidgets.QLineEdit(self)
        project_layout.addWidget(self.project_name_edit, 1, 1)

        table_group = QtWidgets.QGroupBox("Per productie", self)
        table_layout = QtWidgets.QVBoxLayout(table_group)
        main_layout.addWidget(table_group, 1)

        self.table = QtWidgets.QTableWidget(len(productions), 5, self)
        self.table.setHorizontalHeaderLabels([
            "Productie",
            "Leverancier",
            "Documenttype",
            "Nr.",
            "Leveradres",
        ])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        table_layout.addWidget(self.table)

        delivery_opts = [
            "Geen",
            "Bestelling wordt opgehaald",
            "Leveradres wordt nog meegedeeld",
        ] + [
            delivery_db.display_name(a)
            for a in delivery_db.addresses_sorted()
        ]

        doc_type_opts = ["Geen", "Bestelbon", "Offerteaanvraag"]
        self.doc_type_prefixes = {_prefix_for_doc_type(t) for t in doc_type_opts}

        base_options = self._supplier_display_list()

        for row, production in enumerate(productions):
            prod_item = QtWidgets.QTableWidgetItem(production)
            self.table.setItem(row, 0, prod_item)

            supplier_combo = QtWidgets.QComboBox(self)
            supplier_combo.setEditable(True)
            supplier_combo.setInsertPolicy(QtWidgets.QComboBox.InsertPolicy.NoInsert)
            supplier_combo.setFocusPolicy(QtCore.Qt.FocusPolicy.StrongFocus)
            supplier_combo.setMinimumContentsLength(20)
            supplier_combo.setSizeAdjustPolicy(QtWidgets.QComboBox.SizeAdjustPolicy.AdjustToContents)
            supplier_combo.addItems(base_options)
            supplier_combo.lineEdit().editingFinished.connect(self._update_preview)
            supplier_combo.currentTextChanged.connect(self._on_supplier_change)
            self.table.setCellWidget(row, 1, supplier_combo)
            self.sel_combos[production] = supplier_combo

            doc_combo = QtWidgets.QComboBox(self)
            doc_combo.addItems(doc_type_opts)
            doc_combo.currentTextChanged.connect(lambda _text, p=production: self._on_doc_type_change(p))
            self.table.setCellWidget(row, 2, doc_combo)
            self.doc_combos[production] = doc_combo

            doc_edit = QtWidgets.QLineEdit(self)
            self.table.setCellWidget(row, 3, doc_edit)
            self.doc_num_edits[production] = doc_edit

            delivery_combo = QtWidgets.QComboBox(self)
            delivery_combo.addItems(delivery_opts)
            self.table.setCellWidget(row, 4, delivery_combo)
            self.delivery_combos[production] = delivery_combo

        self.table.resizeColumnsToContents()

        preview_group = QtWidgets.QGroupBox("Leverancier details", self)
        preview_layout = QtWidgets.QVBoxLayout(preview_group)
        self.preview_text = QtWidgets.QTextEdit(self)
        self.preview_text.setReadOnly(True)
        preview_layout.addWidget(self.preview_text)
        main_layout.addWidget(preview_group)

        button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Cancel | QtWidgets.QDialogButtonBox.StandardButton.Ok,
            parent=self,
        )
        self.remember_check = QtWidgets.QCheckBox("Onthoud keuze per productie", self)
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.remember_check)
        button_layout.addStretch(1)
        button_layout.addWidget(button_box)
        main_layout.addLayout(button_layout)

        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        self._apply_defaults(initial=True)
        self._update_preview()

    def _supplier_display_list(self):
        suppliers = self.db.suppliers_sorted()
        options = [self.db.display_name(s) for s in suppliers]
        options.insert(0, "(geen)")
        return options

    def _apply_defaults(self, initial=False):
        options = self._supplier_display_list()
        disp_to_name = {self.db.display_name(s): s.supplier for s in self.db.suppliers_sorted()}
        for production, combo in self.sel_combos.items():
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(options)
            default = self.db.get_default(production)
            if default:
                match = next((disp for disp, name in disp_to_name.items() if name.lower() == default.lower()), None)
                if match:
                    combo.setCurrentText(match)
                else:
                    combo.setCurrentIndex(0)
            elif initial and len(options) > 1:
                combo.setCurrentIndex(1)
            else:
                combo.setCurrentIndex(0)
            combo.blockSignals(False)
            self._on_supplier_change(combo.currentText())
        for combo in self.doc_combos.values():
            combo.setCurrentText("Bestelbon")
        for edit in self.doc_num_edits.values():
            edit.clear()

    def _resolve_supplier(self, text):
        if not text:
            return None
        cleaned = text.replace("★ ", "", 1).strip()
        for supplier in self.db.suppliers:
            if supplier.supplier.lower() == cleaned.lower():
                return supplier
        for supplier in self.db.suppliers:
            if cleaned.lower() in supplier.supplier.lower():
                return supplier
        return None

    def _update_preview(self):
        combo = self.sender()
        text = None
        if isinstance(combo, QtWidgets.QComboBox):
            text = combo.currentText()
        elif isinstance(combo, QtWidgets.QLineEdit):
            text = combo.text()
        if not text and combo is not None:
            parent_combo = combo.parent()
            if isinstance(parent_combo, QtWidgets.QComboBox):
                text = parent_combo.currentText()
        supplier = self._resolve_supplier(text or "")
        if not supplier:
            self.preview_text.clear()
            return
        info = [supplier.supplier]
        if supplier.description:
            info.append(supplier.description)
        addr_lines = [line for line in [supplier.adres_1, supplier.adres_2] if line]
        if addr_lines:
            info.append("\n".join(addr_lines))
        for label, value in (
            ("BTW", supplier.btw),
            ("E-mail", supplier.sales_email),
            ("Tel", supplier.phone),
            ("Contact", supplier.contact_sales),
        ):
            if value:
                info.append(f"{label}: {value}")
        self.preview_text.setPlainText("\n".join(info))

    def _on_supplier_change(self, _text):
        sender = self.sender()
        if isinstance(sender, QtWidgets.QComboBox):
            self._update_preview()
            prod = None
            for key, combo in self.sel_combos.items():
                if combo is sender:
                    prod = key
                    break
            if prod:
                doc_combo = self.doc_combos.get(prod)
                doc_edit = self.doc_num_edits.get(prod)
                if doc_combo and doc_edit:
                    if sender.currentText().strip().lower() in ("(geen)", "geen"):
                        doc_combo.setCurrentText("Geen")
                        doc_edit.clear()
                    elif doc_combo.currentText() == "Geen":
                        doc_combo.setCurrentText("Bestelbon")
                        self._on_doc_type_change(prod)

    def _on_doc_type_change(self, production):
        combo = self.doc_combos.get(production)
        edit = self.doc_num_edits.get(production)
        if not combo or not edit:
            return
        prefix = _prefix_for_doc_type(combo.currentText())
        current = edit.text().strip()
        if not current or current in self.doc_type_prefixes:
            edit.setText(prefix)

    def selections(self):
        sel_map = {}
        doc_map = {}
        doc_num_map = {}
        delivery_map = {}
        for production in self.productions:
            supplier_text = self.sel_combos[production].currentText().strip()
            if supplier_text.lower() in ("", "(geen)", "geen"):
                sel_map[production] = ""
            else:
                supplier = self._resolve_supplier(supplier_text)
                sel_map[production] = supplier.supplier if supplier else supplier_text
            doc_map[production] = self.doc_combos[production].currentText()
            doc_num_map[production] = self.doc_num_edits[production].text().strip()
            delivery_map[production] = self.delivery_combos[production].currentText()
        return sel_map, doc_map, doc_num_map, delivery_map

    def project_info(self):
        return self.project_number_edit.text().strip(), self.project_name_edit.text().strip()

    def remember_choice(self):
        return self.remember_check.isChecked()


class CustomBomTab(QtWidgets.QWidget):
    headers = [
        "PartNumber",
        "Description",
        "Production",
        "Bestanden gevonden",
        "Status",
        "Materiaal",
        "Aantal",
        "Oppervlakte",
        "Gewicht",
    ]

    def __init__(self, on_use_callback, parent=None):
        super().__init__(parent)
        self.on_use_callback = on_use_callback
        layout = QtWidgets.QVBoxLayout(self)

        instructions = QtWidgets.QLabel(
            "Plak gegevens uit het klembord (tab-gescheiden). Gebruik Delete om rijen te verwijderen.",
            self,
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        self.table = QtWidgets.QTableView(self)
        self.model = QtGui.QStandardItemModel(0, len(self.headers), self)
        self.model.setHorizontalHeaderLabels(self.headers)
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.installEventFilter(self)
        layout.addWidget(self.table, 1)

        btn_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(btn_layout)
        clear_btn = QtWidgets.QPushButton("Leegmaken", self)
        clear_btn.clicked.connect(self.clear_model)
        btn_layout.addWidget(clear_btn)
        use_btn = QtWidgets.QPushButton("Gebruik als BOM", self)
        use_btn.clicked.connect(self.emit_bom)
        btn_layout.addWidget(use_btn)
        btn_layout.addStretch(1)

    def eventFilter(self, source, event):
        if source is self.table and event.type() == QtCore.QEvent.Type.KeyPress:
            if event.matches(QtGui.QKeySequence.StandardKey.Paste):
                self.paste_from_clipboard()
                return True
            if event.key() in (QtCore.Qt.Key.Key_Delete, QtCore.Qt.Key.Key_Backspace):
                self.delete_selected_rows()
                return True
        return super().eventFilter(source, event)

    def paste_from_clipboard(self):
        clipboard = QtWidgets.QApplication.clipboard()
        text = clipboard.text()
        if not text:
            return
        rows = [row for row in text.splitlines() if row.strip()]
        for row_text in rows:
            parts = [part.strip() for part in row_text.split("\t")]
            row_items = []
            for idx, header in enumerate(self.headers):
                val = parts[idx] if idx < len(parts) else ""
                item = QtGui.QStandardItem(val)
                row_items.append(item)
            self.model.appendRow(row_items)

    def delete_selected_rows(self):
        selection = self.table.selectionModel().selectedRows()
        for index in sorted(selection, key=lambda idx: idx.row(), reverse=True):
            self.model.removeRow(index.row())

    def clear_model(self):
        if QtWidgets.QMessageBox.question(
            self,
            "Bevestigen",
            "Weet je zeker dat je de tabel wilt leegmaken?",
        ) == QtWidgets.QMessageBox.StandardButton.Yes:
            self.model.removeRows(0, self.model.rowCount())

    def emit_bom(self):
        if self.on_use_callback:
            self.on_use_callback(self.to_dataframe())

    def to_dataframe(self):
        import pandas as pd

        data = {header: [] for header in self.headers}
        for row in range(self.model.rowCount()):
            for col, header in enumerate(self.headers):
                index = self.model.index(row, col)
                data[header].append(self.model.data(index) or "")
        return pd.DataFrame(data)

    def from_dataframe(self, df):
        self.model.removeRows(0, self.model.rowCount())
        for _, row in df.iterrows():
            items = []
            for header in self.headers:
                item = QtGui.QStandardItem(str(row.get(header, "")))
                items.append(item)
            self.model.appendRow(items)


class MainWindow(QtWidgets.QMainWindow):
    status_changed = QtCore.pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Filehopper")
        self.resize(1100, 800)

        from suppliers_db import SuppliersDB, SUPPLIERS_DB_FILE
        from clients_db import ClientsDB, CLIENTS_DB_FILE
        from delivery_addresses_db import DeliveryAddressesDB, DELIVERY_DB_FILE

        self.db = SuppliersDB.load(SUPPLIERS_DB_FILE)
        self.client_db = ClientsDB.load(CLIENTS_DB_FILE)
        self.delivery_db = DeliveryAddressesDB.load(DELIVERY_DB_FILE)

        self.source_folder = ""
        self.dest_folder = ""
        self.bom_df = None
        self.item_links = {}

        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        layout = QtWidgets.QVBoxLayout(central)

        self.tabs = QtWidgets.QTabWidget(self)
        layout.addWidget(self.tabs)

        self.main_tab = QtWidgets.QWidget(self)
        self.tabs.addTab(self.main_tab, "Main")

        self.clients_tab = ClientsManagerWidget(self.client_db, on_change=self._on_db_change)
        self.tabs.addTab(self.clients_tab, "Klant beheer")
        self.delivery_tab = DeliveryAddressesManagerWidget(self.delivery_db, on_change=self._on_db_change)
        self.tabs.addTab(self.delivery_tab, "Leveradres beheer")
        self.suppliers_tab = SuppliersManagerWidget(self.db, on_change=self._on_db_change)
        self.tabs.addTab(self.suppliers_tab, "Leverancier beheer")

        self.custom_bom_tab = CustomBomTab(self._use_custom_bom)
        self.tabs.addTab(self.custom_bom_tab, "Custom BOM")

        self._build_main_tab()
        self.status_changed.connect(self._apply_status)

    def _build_main_tab(self):
        layout = QtWidgets.QVBoxLayout(self.main_tab)

        path_group = QtWidgets.QGridLayout()
        layout.addLayout(path_group)
        path_group.addWidget(QtWidgets.QLabel("Bronmap:"), 0, 0)
        self.source_edit = QtWidgets.QLineEdit(self)
        path_group.addWidget(self.source_edit, 0, 1)
        browse_src = QtWidgets.QPushButton("Bladeren", self)
        browse_src.clicked.connect(self._pick_src)
        path_group.addWidget(browse_src, 0, 2)

        path_group.addWidget(QtWidgets.QLabel("Bestemmingsmap:"), 1, 0)
        self.dest_edit = QtWidgets.QLineEdit(self)
        path_group.addWidget(self.dest_edit, 1, 1)
        browse_dst = QtWidgets.QPushButton("Bladeren", self)
        browse_dst.clicked.connect(self._pick_dst)
        path_group.addWidget(browse_dst, 1, 2)

        path_group.addWidget(QtWidgets.QLabel("Opdrachtgever:"), 2, 0)
        self.client_combo = QtWidgets.QComboBox(self)
        self._refresh_clients_combo()
        path_group.addWidget(self.client_combo, 2, 1)
        manage_btn = QtWidgets.QPushButton("Beheer", self)
        manage_btn.clicked.connect(lambda: self.tabs.setCurrentWidget(self.clients_tab))
        path_group.addWidget(manage_btn, 2, 2)

        filter_group = QtWidgets.QGroupBox("Selecteer bestandstypen om te kopiëren", self)
        filter_layout = QtWidgets.QHBoxLayout(filter_group)
        layout.addWidget(filter_group)
        self.pdf_check = QtWidgets.QCheckBox("PDF (.pdf)", self)
        self.step_check = QtWidgets.QCheckBox("STEP (.step, .stp)", self)
        self.dxf_check = QtWidgets.QCheckBox("DXF (.dxf)", self)
        self.dwg_check = QtWidgets.QCheckBox("DWG (.dwg)", self)
        filter_layout.addWidget(self.pdf_check)
        filter_layout.addWidget(self.step_check)
        filter_layout.addWidget(self.dxf_check)
        filter_layout.addWidget(self.dwg_check)
        filter_layout.addStretch(1)

        toolbar_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(toolbar_layout)
        load_bom_btn = QtWidgets.QPushButton("Laad BOM (CSV/Excel)", self)
        load_bom_btn.clicked.connect(self._load_bom)
        toolbar_layout.addWidget(load_bom_btn)
        check_files_btn = QtWidgets.QPushButton("Controleer Bestanden", self)
        check_files_btn.clicked.connect(self._check_files)
        toolbar_layout.addWidget(check_files_btn)
        copy_flat_btn = QtWidgets.QPushButton("Kopieer zonder submappen", self)
        copy_flat_btn.clicked.connect(self._copy_flat)
        toolbar_layout.addWidget(copy_flat_btn)
        copy_per_btn = QtWidgets.QPushButton("Kopieer per productie + bestelbonnen", self)
        copy_per_btn.clicked.connect(self._copy_per_prod)
        toolbar_layout.addWidget(copy_per_btn)
        combine_btn = QtWidgets.QPushButton("Combine pdf", self)
        combine_btn.clicked.connect(self._combine_pdf)
        toolbar_layout.addWidget(combine_btn)
        custom_btn = QtWidgets.QPushButton("Custom BOM", self)
        custom_btn.clicked.connect(lambda: self.tabs.setCurrentWidget(self.custom_bom_tab))
        toolbar_layout.addWidget(custom_btn)
        toolbar_layout.addStretch(1)

        manual_group = QtWidgets.QGroupBox("PartNumbers (één per lijn)", self)
        manual_layout = QtWidgets.QVBoxLayout(manual_group)
        layout.addWidget(manual_group)
        self.manual_text = QtWidgets.QPlainTextEdit(self)
        manual_layout.addWidget(self.manual_text)
        use_manual_btn = QtWidgets.QPushButton("Gebruik PartNumbers", self)
        use_manual_btn.clicked.connect(self._load_manual_pns)
        manual_layout.addWidget(use_manual_btn, alignment=QtCore.Qt.AlignmentFlag.AlignLeft)

        self.bom_table = QtWidgets.QTableWidget(0, 5, self)
        self.bom_table.setHorizontalHeaderLabels([
            "PartNumber",
            "Description",
            "Production",
            "Bestanden gevonden",
            "Status",
        ])
        self.bom_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.bom_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.bom_table.cellClicked.connect(self._on_table_click)
        layout.addWidget(self.bom_table, 1)

        actions_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(actions_layout)
        self.zip_check = QtWidgets.QCheckBox("Zip per productie", self)
        actions_layout.addWidget(self.zip_check)
        actions_layout.addStretch(1)

        self.status_label = QtWidgets.QLabel("Klaar", self)
        layout.addWidget(self.status_label)

    def _on_db_change(self):
        self._refresh_clients_combo()

    def _refresh_clients_combo(self):
        current = self.client_combo.currentText() if hasattr(self, "client_combo") else ""
        options = [self.client_db.display_name(c) for c in self.client_db.clients_sorted()]
        if hasattr(self, "client_combo"):
            self.client_combo.clear()
            self.client_combo.addItems(options)
            if current in options:
                self.client_combo.setCurrentText(current)
            elif options:
                self.client_combo.setCurrentIndex(0)

    def _pick_src(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Bronmap selecteren")
        if path:
            self.source_folder = path
            self.source_edit.setText(path)

    def _pick_dst(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Bestemmingsmap selecteren")
        if path:
            self.dest_folder = path
            self.dest_edit.setText(path)

    def _selected_exts(self):
        exts = []
        if self.pdf_check.isChecked():
            exts.append(".pdf")
        if self.step_check.isChecked():
            exts.extend([".step", ".stp"])
        if self.dxf_check.isChecked():
            exts.append(".dxf")
        if self.dwg_check.isChecked():
            exts.append(".dwg")
        return exts or None

    def _load_bom(self):
        start_dir = self.source_folder or ""
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Selecteer BOM",
            start_dir,
            "CSV (*.csv);;Excel (*.xlsx *.xls)",
        )
        if not path:
            return
        try:
            from bom import load_bom

            df = load_bom(path)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Fout", str(exc))
            return
        if "Bestanden gevonden" not in df.columns:
            df["Bestanden gevonden"] = ""
        if "Status" not in df.columns:
            df["Status"] = ""
        self.bom_df = df
        self._refresh_bom_table()
        self.status_changed.emit(f"BOM geladen: {len(df)} rijen")

    def _load_manual_pns(self):
        text = self.manual_text.toPlainText().strip()
        if not text:
            return
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if not lines:
            return
        import pandas as pd

        data = {
            "PartNumber": lines,
            "Description": ["" for _ in lines],
            "Production": ["" for _ in lines],
            "Bestanden gevonden": ["" for _ in lines],
            "Status": ["" for _ in lines],
            "Materiaal": ["" for _ in lines],
            "Aantal": [1 for _ in lines],
            "Oppervlakte": ["" for _ in lines],
            "Gewicht": ["" for _ in lines],
        }
        self.bom_df = pd.DataFrame(data)
        self._refresh_bom_table()
        self.status_changed.emit(f"Partnummers geladen: {len(lines)} rijen")

    def _refresh_bom_table(self):
        self.item_links.clear()
        self.bom_table.setRowCount(0)
        if self.bom_df is None:
            return
        for _, row in self.bom_df.iterrows():
            table_row = self.bom_table.rowCount()
            self.bom_table.insertRow(table_row)
            values = [
                row.get("PartNumber", ""),
                row.get("Description", ""),
                row.get("Production", ""),
                row.get("Bestanden gevonden", ""),
                row.get("Status", ""),
            ]
            for col, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(str(value))
                if col == 4:
                    item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                self.bom_table.setItem(table_row, col, item)
            link = row.get("Link")
            if link:
                self.item_links[table_row] = link
        self.bom_table.resizeColumnsToContents()

    def _on_table_click(self, row, column):
        if column != 4:
            return
        if self.bom_table.item(row, column).text() != "❌":
            return
        path = self.item_links.get(row)
        if not path or not os.path.exists(path):
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception:
            pass

    def _check_files(self):
        if self.bom_df is None:
            QtWidgets.QMessageBox.warning(self, "Let op", "Laad eerst een BOM.")
            return
        if not self.source_folder:
            QtWidgets.QMessageBox.warning(self, "Let op", "Selecteer een bronmap.")
            return
        exts = self._selected_exts()
        if not exts:
            QtWidgets.QMessageBox.warning(self, "Let op", "Selecteer minstens één bestandstype.")
            return
        self.status_changed.emit("Bezig met controleren...")

        def work():
            from helpers import _build_file_index

            idx = _build_file_index(self.source_folder, exts)
            sw_idx = _build_file_index(self.source_folder, [".sldprt", ".slddrw"])
            found = []
            status = []
            links = []
            groups = []
            exts_set = {ext.lower() for ext in exts}
            if ".step" in exts_set or ".stp" in exts_set:
                groups.append({".step", ".stp"})
                exts_set -= {".step", ".stp"}
            for ext in exts_set:
                groups.append({ext})
            for _, row in self.bom_df.iterrows():
                pn = row.get("PartNumber")
                hits = idx.get(pn, [])
                hit_exts = {os.path.splitext(h)[1].lower() for h in hits}
                all_present = all(any(ext in hit_exts for ext in group) for group in groups)
                found.append(", ".join(sorted(e.lstrip('.') for e in hit_exts)))
                status.append("✅" if all_present else "❌")
                link = ""
                if not all_present:
                    missing = []
                    for group in groups:
                        if not any(ext in hit_exts for ext in group):
                            missing.extend(group)
                    sw_hits = sw_idx.get(pn, [])
                    drw = next((p for p in sw_hits if p.lower().endswith(".slddrw")), None)
                    prt = next((p for p in sw_hits if p.lower().endswith(".sldprt")), None)
                    if ".pdf" in missing and drw:
                        link = drw
                    elif prt:
                        link = prt
                    elif drw:
                        link = drw
                links.append(link)
            def update():
                self.bom_df["Bestanden gevonden"] = found
                self.bom_df["Status"] = status
                self.bom_df["Link"] = links
                self._refresh_bom_table()
                self.status_changed.emit("Controle klaar.")
            QtCore.QTimer.singleShot(0, update)

        threading.Thread(target=work, daemon=True).start()

    def _copy_flat(self):
        exts = self._selected_exts()
        if not exts or not self.source_folder or not self.dest_folder:
            QtWidgets.QMessageBox.warning(self, "Let op", "Selecteer bron, bestemming en extensies.")
            return

        def work():
            from helpers import _build_file_index

            self.status_changed.emit("Kopiëren...")
            idx = _build_file_index(self.source_folder, exts)
            cnt = 0
            for _, paths in idx.items():
                for path in paths:
                    dst = os.path.join(self.dest_folder, os.path.basename(path))
                    shutil.copy2(path, dst)
                    cnt += 1
            def done():
                self.status_changed.emit(f"Gekopieerd: {cnt}")
            QtCore.QTimer.singleShot(0, done)

        threading.Thread(target=work, daemon=True).start()

    def _resolve_delivery(self, name):
        clean = name.replace("★ ", "", 1)
        if clean == "Geen":
            return None
        if clean in ("Bestelling wordt opgehaald", "Leveradres wordt nog meegedeeld"):
            from delivery_addresses_db import DeliveryAddress

            return DeliveryAddress(name=clean)
        return self.delivery_db.get(clean)

    def _copy_per_prod(self):
        if self.bom_df is None:
            QtWidgets.QMessageBox.warning(self, "Let op", "Laad eerst een BOM.")
            return
        exts = self._selected_exts()
        if not exts or not self.source_folder or not self.dest_folder:
            QtWidgets.QMessageBox.warning(self, "Let op", "Selecteer bron, bestemming en extensies.")
            return
        productions = sorted({(str(row.get("Production") or "").strip() or "_Onbekend") for _, row in self.bom_df.iterrows()})
        dialog = SupplierSelectionDialog(productions, self.db, self.delivery_db, self)
        if dialog.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            return
        sel_map, doc_map, doc_num_map, delivery_map_raw = dialog.selections()
        project_number, project_name = dialog.project_info()
        remember = dialog.remember_choice()

        def work():
            from orders import copy_per_production_and_orders, DEFAULT_FOOTER_NOTE

            self.status_changed.emit("Kopiëren & bestelbonnen maken...")
            client = self.client_db.get(self.client_combo.currentText().replace("★ ", "", 1))
            resolved_delivery = {prod: self._resolve_delivery(name) for prod, name in delivery_map_raw.items()}
            cnt, chosen = copy_per_production_and_orders(
                self.source_folder,
                self.dest_folder,
                self.bom_df,
                exts,
                self.db,
                sel_map,
                doc_map,
                doc_num_map,
                remember,
                client=client,
                delivery_map=resolved_delivery,
                footer_note=DEFAULT_FOOTER_NOTE,
                zip_parts=bool(self.zip_check.isChecked()),
                project_number=project_number,
                project_name=project_name,
            )

            def done():
                self.status_changed.emit(f"Klaar. Gekopieerd: {cnt}. Leveranciers: {chosen}")
                QtWidgets.QMessageBox.information(self, "Klaar", "Bestelbonnen aangemaakt.")
                if sys.platform.startswith("win"):
                    try:
                        os.startfile(self.dest_folder)
                    except Exception:
                        pass

            QtCore.QTimer.singleShot(0, done)

        threading.Thread(target=work, daemon=True).start()

    def _combine_pdf(self):
        if self.source_folder and self.bom_df is not None:
            def work_source():
                from orders import combine_pdfs_from_source

                self.status_changed.emit("PDF's combineren...")
                try:
                    out_dir = self.dest_folder or self.source_folder
                    cnt = combine_pdfs_from_source(self.source_folder, self.bom_df, out_dir)
                except ModuleNotFoundError:
                    def missing():
                        self.status_changed.emit("PyPDF2 ontbreekt")
                        QtWidgets.QMessageBox.warning(self, "PyPDF2 ontbreekt", "Installeer PyPDF2 om PDF's te combineren.")
                    QtCore.QTimer.singleShot(0, missing)
                    return
                def done():
                    self.status_changed.emit(f"Gecombineerde pdf's: {cnt}")
                    QtWidgets.QMessageBox.information(self, "Klaar", "PDF's gecombineerd.")
                QtCore.QTimer.singleShot(0, done)
            threading.Thread(target=work_source, daemon=True).start()
        elif self.dest_folder:
            def work_dest():
                from orders import combine_pdfs_per_production

                self.status_changed.emit("PDF's combineren...")
                try:
                    cnt = combine_pdfs_per_production(self.dest_folder)
                except ModuleNotFoundError:
                    def missing():
                        self.status_changed.emit("PyPDF2 ontbreekt")
                        QtWidgets.QMessageBox.warning(self, "PyPDF2 ontbreekt", "Installeer PyPDF2 om PDF's te combineren.")
                    QtCore.QTimer.singleShot(0, missing)
                    return
                def done():
                    self.status_changed.emit(f"Gecombineerde pdf's: {cnt}")
                    QtWidgets.QMessageBox.information(self, "Klaar", "PDF's gecombineerd.")
                QtCore.QTimer.singleShot(0, done)
            threading.Thread(target=work_dest, daemon=True).start()
        else:
            QtWidgets.QMessageBox.warning(self, "Let op", "Selecteer bron + BOM of bestemmingsmap.")

    def _apply_status(self, text):
        self.status_label.setText(text)

    def _use_custom_bom(self, df):
        if df is None:
            return
        if "Bestanden gevonden" not in df.columns:
            df["Bestanden gevonden"] = ""
        if "Status" not in df.columns:
            df["Status"] = ""
        self.bom_df = df
        self._refresh_bom_table()
        self.tabs.setCurrentWidget(self.main_tab)
        self.status_changed.emit(f"Custom BOM geladen: {len(df)} rijen")


def start_gui():

    app = QtWidgets.QApplication.instance()
    owns_app = False
    if app is None:
        owns_app = True
        app = QtWidgets.QApplication(sys.argv or [""])
    window = MainWindow()
    window.show()
    if owns_app:
        app.exec()
