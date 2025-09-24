from app_settings import AppSettings, FileExtensionSetting, DEFAULT_FILE_EXTENSIONS


def test_app_settings_roundtrip(tmp_path):
    path = tmp_path / "app_settings.json"
    settings = AppSettings(
        source_folder="/tmp/src",
        dest_folder="/tmp/dst",
        project_number="PN-123",
        project_name="Demo",
        file_extensions=[
            FileExtensionSetting(
                key="pdf",
                label="PDF (.pdf)",
                patterns=[".pdf"],
                enabled=True,
            ),
            FileExtensionSetting(
                key="dwg",
                label="DWG (.dwg)",
                patterns=[".dwg"],
                enabled=False,
            ),
        ],
        zip_per_production=False,
        export_date_prefix=True,
        export_date_suffix=False,
        custom_prefix_enabled=True,
        custom_prefix_text="PRE",
        custom_suffix_enabled=True,
        custom_suffix_text="SUF",
        bundle_latest=True,
        bundle_dry_run=True,
    )
    settings.save(path)

    loaded = AppSettings.load(path)

    assert loaded == settings


def test_app_settings_migrates_legacy_booleans(tmp_path):
    payload = {
        "source_folder": "/src",
        "dest_folder": "/dst",
        "project_number": "PN",
        "project_name": "Demo",
        "pdf": True,
        "step": False,
        "dxf": True,
        "dwg": False,
    }
    path = tmp_path / "legacy.json"
    path.write_text("", encoding="utf-8")
    loaded = AppSettings.from_dict(payload)
    assert [ext.key for ext in loaded.file_extensions] == [
        ext.key for ext in DEFAULT_FILE_EXTENSIONS
    ]
    status = {ext.key: ext.enabled for ext in loaded.file_extensions}
    assert status["pdf"] is True
    assert status["dxf"] is True
    assert status["step"] is False
    assert status["dwg"] is False


def test_app_settings_corrupt_file_returns_defaults(tmp_path):
    path = tmp_path / "app_settings.json"
    path.write_text("{not valid json", encoding="utf-8")

    loaded = AppSettings.load(path)

    assert loaded == AppSettings()
