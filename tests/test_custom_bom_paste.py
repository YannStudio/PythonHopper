import ast
import pathlib


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
