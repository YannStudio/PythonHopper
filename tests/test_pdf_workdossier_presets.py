from pdf_workdossier_presets import (
    PdfWorkDossierPreset,
    PdfWorkDossierPresetsDB,
    PdfWorkDossierSection,
    default_pdf_workdossier_preset,
)


def test_pdf_workdossier_preset_roundtrip(tmp_path):
    path = tmp_path / "pdf_workdossier_presets.json"
    db = PdfWorkDossierPresetsDB(
        [
            PdfWorkDossierPreset(
                name="Werkdossier klant",
                sections=[
                    PdfWorkDossierSection(
                        "Hoofdassembly",
                        include_bom_pdf=True,
                    ),
                    PdfWorkDossierSection(
                        "Laserwerk",
                        identifiers=["Laser", "Lasersnijden"],
                    ),
                ],
            )
        ]
    )

    db.save(str(path))
    loaded = PdfWorkDossierPresetsDB.load(str(path))

    preset = loaded.get("Werkdossier klant")
    assert preset is not None
    assert preset.sections[0].include_bom_pdf is True
    assert preset.sections[1].identifiers == ["Laser", "Lasersnijden"]


def test_default_pdf_workdossier_preset_contains_expected_sections():
    preset = default_pdf_workdossier_preset()

    names = [section.name for section in preset.sections]
    assert "Hoofdassembly" in names
    assert "Assembly tekeningen" in names
    assert "Tube laserwerk" in names
    assert "Spare parts" in names
