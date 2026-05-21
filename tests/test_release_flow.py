import build_executable
from scripts import release


def test_windows_version_text_pads_to_four_parts():
    assert release.windows_version_text("3.1") == "3.1.0.0"
    assert release.windows_version_text("4.2.5") == "4.2.5.0"


def test_release_dist_dir_uses_version():
    assert build_executable.release_dist_dir("3.1").name == "Filehopper-3.1"


def test_pyinstaller_cmd_uses_custom_dist_dir(tmp_path, monkeypatch):
    monkeypatch.setattr(build_executable.platform, "system", lambda: "Linux")

    cmd = build_executable._pyinstaller_cmd(
        "main.py",
        "filehopper-test",
        windowed=True,
        onefile=False,
        data_files=[],
        dist_dir=tmp_path,
    )

    assert "--distpath" in cmd
    assert cmd[cmd.index("--distpath") + 1] == str(tmp_path)
