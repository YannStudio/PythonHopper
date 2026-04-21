import ast
import pathlib


def test_add_supplier_uses_full_edit_dialog_instead_of_name_prompt():
    source = pathlib.Path("gui.py").read_text(encoding="utf-8")
    mod = ast.parse(source)
    start = next(
        node
        for node in mod.body
        if isinstance(node, ast.FunctionDef) and node.name == "start_gui"
    )
    manager_cls = next(
        node
        for node in start.body
        if isinstance(node, ast.ClassDef) and node.name == "SuppliersManagerFrame"
    )
    add_supplier = next(
        node
        for node in manager_cls.body
        if isinstance(node, ast.FunctionDef) and node.name == "add_supplier"
    )

    assert not any(
        isinstance(node, ast.Attribute) and node.attr == "askstring"
        for node in ast.walk(add_supplier)
    )
    assert any(
        isinstance(node, ast.Attribute) and node.attr == "_EditDialog"
        for node in ast.walk(add_supplier)
    )


def test_supplier_edit_dialog_supports_custom_title():
    source = pathlib.Path("gui.py").read_text(encoding="utf-8")
    mod = ast.parse(source)
    start = next(
        node
        for node in mod.body
        if isinstance(node, ast.FunctionDef) and node.name == "start_gui"
    )
    manager_cls = next(
        node
        for node in start.body
        if isinstance(node, ast.ClassDef) and node.name == "SuppliersManagerFrame"
    )
    edit_dialog = next(
        node
        for node in manager_cls.body
        if isinstance(node, ast.ClassDef) and node.name == "_EditDialog"
    )
    init_method = next(
        node
        for node in edit_dialog.body
        if isinstance(node, ast.FunctionDef) and node.name == "__init__"
    )

    arg_names = [arg.arg for arg in init_method.args.args]
    kwonly_names = [arg.arg for arg in init_method.args.kwonlyargs]

    assert arg_names[:3] == ["self", "master", "supplier"]
    assert "title" in kwonly_names
