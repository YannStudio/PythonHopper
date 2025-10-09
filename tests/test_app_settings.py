import json

from app_settings import AppSettings, FileExtensionSetting, DEFAULT_FILE_EXTENSIONS
from orders import DEFAULT_FOOTER_NOTE


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
                key="igs",
                label="IGES (.iges, .igs)",
                patterns=[".iges", ".igs"],
                enabled=False,
            ),
        ],
        zip_per_production=False,
        copy_finish_exports=True,
        zip_finish_exports=False,
        export_processed_bom=False,
        export_date_prefix=True,
        export_date_suffix=False,
        custom_prefix_enabled=True,
        custom_prefix_text="PRE",
        custom_suffix_enabled=True,
        custom_suffix_text="SUF",
        bundle_latest=True,
        bundle_dry_run=True,
        footer_note="Aangepaste voorwaarden",
    )
    settings.save(path)

    loaded = AppSettings.load(path)

    assert loaded == settings


def test_app_settings_corrupt_file_returns_defaults(tmp_path):
    path = tmp_path / "app_settings.json"
    path.write_text("{not valid json", encoding="utf-8")

    loaded = AppSettings.load(path)

    assert loaded == AppSettings()


def test_app_settings_loads_legacy_extension_flags(tmp_path):
    path = tmp_path / "app_settings.json"
    payload = {
        "source_folder": "/tmp/src",
        "pdf": 1,
        "step": False,
        "dxf": "true",
    }
    path.write_text(json.dumps(payload), encoding="utf-8")

    loaded = AppSettings.load(path)

    assert any(ext.key == "pdf" and ext.enabled for ext in loaded.file_extensions)
    assert any(ext.key == "dxf" and ext.enabled for ext in loaded.file_extensions)
    assert any(ext.key == "step" and not ext.enabled for ext in loaded.file_extensions)
    loaded_keys = {ext.key for ext in loaded.file_extensions}
    assert {ext.key for ext in DEFAULT_FILE_EXTENSIONS}.issubset(loaded_keys)


def test_default_footer_note_matches_orders_constant():
    assert AppSettings().footer_note == DEFAULT_FOOTER_NOTE
