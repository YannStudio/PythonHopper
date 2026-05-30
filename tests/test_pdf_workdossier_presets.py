from pdf_workdossier_presets import (
    PdfWorkDossierPreset,
    PdfWorkDossierPresetsDB,
    PdfWorkDossierSection,
    default_pdf_workdossier_preset,
    tecno_art_pdf_workdossier_preset,
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
                    PdfWorkDossierSection(
                        "Overige",
                        include_unmatched=True,
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
    assert preset.sections[2].include_unmatched is True


def test_default_pdf_workdossier_preset_contains_expected_sections():
    preset = default_pdf_workdossier_preset()

    names = [section.name for section in preset.sections]
    assert "Hoofdassembly" in names
    assert "Assembly tekeningen" in names
    assert "Tube laserwerk" in names
    assert "Spare parts" in names
    assert "Overige" in names


def test_tecno_art_pdf_workdossier_preset_contains_requested_order():
    preset = tecno_art_pdf_workdossier_preset()

    assert [section.name for section in preset.sections] == [
        "Hoofdassembly",
        "Assembly",
        "Weld assembly",
        "Mount material",
        "Spare parts",
        "Cutting",
        "Lasercutting",
        "Laser cutting +4m",
        "Tube laser",
        "Tube laser L",
        "Overige producties",
    ]
    assert preset.sections[-1].include_unmatched is True
