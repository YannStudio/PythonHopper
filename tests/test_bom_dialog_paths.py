import os

from gui import _resolve_file_dialog_initial_dir


def test_file_dialog_initial_dir_uses_existing_directory(tmp_path):
    source_dir = tmp_path / "drive-ready"
    source_dir.mkdir()

    assert _resolve_file_dialog_initial_dir(str(source_dir)) == str(source_dir.resolve())


def test_file_dialog_initial_dir_falls_back_when_directory_is_missing(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    assert _resolve_file_dialog_initial_dir(str(tmp_path / "missing-drive")) == os.getcwd()

