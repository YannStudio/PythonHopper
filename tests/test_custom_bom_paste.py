import ast
import pathlib
import types

import pandas as pd
import pytest


def test_custom_bom_frame_paste_bindings():
    source = pathlib.Path("gui.py").read_text()
    mod = ast.parse(source)
    start = next(
        node for node in mod.body if isinstance(node, ast.FunctionDef) and node.name == "start_gui"
    )
    custom_cls = next(
        node for node in start.body if isinstance(node, ast.ClassDef) and node.name == "CustomBOMFrame"
    )
    init_fn = next(
        node for node in custom_cls.body if isinstance(node, ast.FunctionDef) and node.name == "__init__"
    )
    events = []
    for node in ast.walk(init_fn):
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
            if node.func.attr == "bind" and node.args and isinstance(node.args[0], ast.Constant):
                events.append(node.args[0].value)
    assert "<Control-v>" in events
    assert "<Command-v>" in events


def _load_method(name, messagebox):
    source = pathlib.Path("gui.py").read_text()
    mod = ast.parse(source)
    start = next(
        node for node in mod.body if isinstance(node, ast.FunctionDef) and node.name == "start_gui"
    )
    custom_cls = next(
        node for node in start.body if isinstance(node, ast.ClassDef) and node.name == "CustomBOMFrame"
    )
    method = next(
        node for node in custom_cls.body if isinstance(node, ast.FunctionDef) and node.name == name
    )
    module = ast.Module(body=[method], type_ignores=[])
    ast.fix_missing_locations(module)
    ns = {}
    exec(compile(module, "gui.py", "exec"), {"pd": pd, "messagebox": messagebox}, ns)
    return ns[name]


class _DummyTree:
    def __init__(self):
        self.rows = []

    def get_children(self):
        return list(range(len(self.rows)))

    def item(self, idx, option=None, values=None):
        if values is not None:
            self.rows[idx] = tuple(values)
        if option == "values":
            return self.rows[idx]
        return {"values": self.rows[idx]}

    def insert(self, _parent, _index, values):
        self.rows.append(tuple(values))
        return len(self.rows) - 1

    def delete(self, idx):
        self.rows.pop(idx)

    def selection(self):
        return []


class _DummyMsgBox:
    def __init__(self):
        self.args = None

    def showerror(self, title, message):
        self.args = (title, message)


class _DummyFrame:
    COLS = (
        "PartNumber",
        "Description",
        "Production",
        "QTY",
        "Material",
    )

    def __init__(self, text="", fail_clip=False):
        self.tree = _DummyTree()
        self._paste_cell = None
        self._text = text
        self._fail_clip = fail_clip
        self.on_save = None

    def _add_row(self):
        self.tree.insert("", "end", values=("", "", "", 1, ""))

    def clipboard_get(self):
        if self._fail_clip:
            raise RuntimeError("no clipboard")
        return self._text


def test_on_paste_row_and_column(monkeypatch):
    msg = _DummyMsgBox()
    fn = _load_method("_on_paste", msg)

    def read_rows(*_a, **_k):
        return pd.DataFrame([
            ["PN1", "D1", "P1", "5", "M1"],
            ["PN2", "D2", "P2", "7", "M2"],
        ])

    monkeypatch.setattr(pd, "read_clipboard", read_rows)
    frame = _DummyFrame()
    frame._on_paste = types.MethodType(fn, frame)
    frame._on_paste()
    assert frame.tree.rows[0] == ("PN1", "D1", "P1", "5", "M1")
    assert frame.tree.rows[1] == ("PN2", "D2", "P2", "7", "M2")

    def read_col(*_a, **_k):
        return pd.DataFrame(["8", "9"])

    monkeypatch.setattr(pd, "read_clipboard", read_col)
    frame = _DummyFrame()
    frame._add_row()
    frame._paste_cell = (0, 3)
    frame._on_paste = types.MethodType(fn, frame)
    frame._on_paste()
    assert frame.tree.rows[0][3] == "8"
    assert frame.tree.rows[1][3] == "9"


def test_on_paste_clipboard_error(monkeypatch):
    msg = _DummyMsgBox()
    fn = _load_method("_on_paste", msg)

    def raise_err(*_a, **_k):
        raise RuntimeError("no clipboard")

    monkeypatch.setattr(pd, "read_clipboard", raise_err)
    frame = _DummyFrame(fail_clip=True)
    frame._on_paste = types.MethodType(fn, frame)
    frame._on_paste()
    assert msg.args is not None
    assert "xclip" in msg.args[1] or "xsel" in msg.args[1]
    assert frame.tree.rows == []


def test_save_validates_qty(monkeypatch):
    msg = _DummyMsgBox()
    fn = _load_method("_save", msg)

    saved = {}
    frame = _DummyFrame()
    frame.tree.rows = [("PN", "D", "P", "X", "M")]
    frame._save = types.MethodType(fn, frame)
    frame.on_save = lambda df: saved.setdefault("df", df)
    frame._save()
    assert msg.args is not None
    assert saved == {}

    msg2 = _DummyMsgBox()
    fn2 = _load_method("_save", msg2)
    frame2 = _DummyFrame()
    frame2.tree.rows = [("PN", "D", "P", "5", "M")]
    frame2._save = types.MethodType(fn2, frame2)
    captured = {}
    frame2.on_save = lambda df: captured.setdefault("df", df)
    frame2._save()
    assert msg2.args is None
    assert list(captured["df"].columns) == list(_DummyFrame.COLS)
    assert captured["df"]["QTY"].tolist() == [5]
