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


def _load_on_paste(messagebox):
    """Return the `_on_paste` function object from gui.CustomBOMFrame."""
    source = pathlib.Path("gui.py").read_text()
    mod = ast.parse(source)
    start = next(
        node for node in mod.body if isinstance(node, ast.FunctionDef) and node.name == "start_gui"
    )
    custom_cls = next(
        node for node in start.body if isinstance(node, ast.ClassDef) and node.name == "CustomBOMFrame"
    )
    on_paste = next(
        node for node in custom_cls.body if isinstance(node, ast.FunctionDef) and node.name == "_on_paste"
    )
    module = ast.Module(body=[on_paste], type_ignores=[])
    ast.fix_missing_locations(module)
    ns = {}
    exec(compile(module, "gui.py", "exec"), {"pd": pd, "messagebox": messagebox}, ns)
    return ns["_on_paste"]


class _DummyTree:
    def __init__(self):
        self.rows = []

    def get_children(self):
        return list(range(len(self.rows)))

    def item(self, idx, option=None, values=None):
        if values is not None:
            self.rows[idx] = tuple(values)
        return {"values": self.rows[idx]}

    def insert(self, _parent, _index, values):
        self.rows.append(tuple(values))
        return len(self.rows) - 1


class _DummyMsgBox:
    def __init__(self):
        self.args = None

    def showerror(self, title, message):
        self.args = (title, message)


class _DummyFrame:
    COLS = (
        "PartNumber",
        "Description",
        "Materiaal",
        "Aantal",
        "Oppervlakte",
        "Gewicht",
    )

    def __init__(self, text="", fail_clip=False):
        self.tree = _DummyTree()
        self._paste_cell = None
        self._text = text
        self._fail_clip = fail_clip

    def _add_row(self):
        self.tree.insert("", "end", values=("", "", "", 1, "", ""))

    def clipboard_get(self):
        if self._fail_clip:
            raise RuntimeError("no clipboard")
        return self._text


def test_on_paste_fallback(monkeypatch):
    """When pandas clipboard read fails, fallback to manual parsing."""
    msg = _DummyMsgBox()
    fn = _load_on_paste(msg)

    def raise_err(*_a, **_k):
        raise RuntimeError("no clipboard")

    monkeypatch.setattr(pd, "read_clipboard", raise_err)

    frame = _DummyFrame("A\tB\nC\tD")
    frame._on_paste = types.MethodType(fn, frame)
    frame._on_paste()

    assert frame.tree.rows[0][:2] == ("A", "B")
    assert frame.tree.rows[1][:2] == ("C", "D")
    assert msg.args is None


def test_on_paste_clipboard_error(monkeypatch):
    """When clipboard access is unavailable, show an error."""
    msg = _DummyMsgBox()
    fn = _load_on_paste(msg)

    def raise_err(*_a, **_k):
        raise RuntimeError("no clipboard")

    monkeypatch.setattr(pd, "read_clipboard", raise_err)

    frame = _DummyFrame(fail_clip=True)
    frame._on_paste = types.MethodType(fn, frame)
    frame._on_paste()

    assert msg.args is not None
    assert "xclip" in msg.args[1] or "xsel" in msg.args[1]
    assert frame.tree.rows == []
