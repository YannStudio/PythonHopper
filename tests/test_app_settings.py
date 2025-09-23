from app_settings import AppSettings


def test_app_settings_roundtrip(tmp_path):
    path = tmp_path / "app_settings.json"
    settings = AppSettings(
        source_folder="/tmp/src",
        dest_folder="/tmp/dst",
        project_number="PN-123",
        project_name="Demo",
        pdf=True,
        step=True,
        dxf=False,
        dwg=True,
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


def test_app_settings_corrupt_file_returns_defaults(tmp_path):
    path = tmp_path / "app_settings.json"
    path.write_text("{not valid json", encoding="utf-8")

    loaded = AppSettings.load(path)

    assert loaded == AppSettings()
