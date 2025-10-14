# Filehopper
Filehopper in Python

## Installatie

De GUI-componenten vertrouwen op [`pandastable`](https://github.com/dmnfarrell/pandastable)
voor het spreadsheet-gedeelte van de custom BOM-tab. Alle runtime
afhankelijkheden worden via `requirements.txt` geïnstalleerd:

```bash
pip install -r requirements.txt
```

Start vervolgens de applicatie met `python main.py` of `python -m gui`.

## Zelfstandige executables bouwen

Je kan zowel voor macOS als Windows een standalone-versie maken met behulp van
PyInstaller. Installeer eerst de afhankelijkheden voor runtime en build:

```bash
pip install -r requirements.txt
pip install -r requirements-build.txt
```

Bouw daarna de gewenste uitvoerbestanden. Op het doelplatform (macOS of
Windows) voer je bijvoorbeeld uit:

```bash
python build_executable.py --target macos
python build_executable.py --target windows
```

Standaard worden er twee varianten gemaakt:

* `filehopper-<target>` — console-app met CLI-ondersteuning.
* `filehopper-gui-<target>` — windowed app die rechtstreeks de GUI opent.

De resultaten verschijnen in de map `dist/`.

### Gebruikersgegevens

Wanneer je een gebundelde versie gebruikt, worden de databestanden per gebruiker
opgeslagen. Op Windows vind je ze in `%LOCALAPPDATA%\Filehopper`, op macOS in
`~/Library/Application Support/Filehopper`. Hierdoor blijven gegevens behouden
na updates.

## Voorbeelden

## Problemen oplossen

### "Please clean your repository working tree before checkout"

Deze melding verschijnt in Visual Studio (Code) wanneer je probeert te
wisselen van branch of commit terwijl er lokale wijzigingen zijn. Los dit op
door eerst je werkdirectory op te ruimen:

1. Controleer de status van de repository:

   ```bash
   git status
   ```

2. Commit of stash de wijzigingen:

   ```bash
   git add .
   git commit -m "Jouw boodschap"
   # of
   git stash
   ```

3. Probeer daarna opnieuw van branch of commit te wisselen.

Als je wijzigingen niet wil bewaren, kan je ze weggooien met `git checkout --
<bestand>` of `git reset --hard`, maar wees hiermee voorzichtig: je verliest
dan alle niet-gecommit werk.

### Leverancier toevoegen

```
python cli.py suppliers add "ACME" \
    --btw BE123456789 \
    --adres-1 "Teststraat 1" \
    --adres-2 "2000 Antwerpen" \
    --tel "+32 123" \
    --email "sales@acme.com"
```

### Leveranciers importeren uit CSV

```
python cli.py suppliers import-csv suppliers.csv \
    --btw BE123456789 \
    --adres-1 "Teststraat 1" \
    --adres-2 "2000 Antwerpen" \
    --tel "+32 123" \
    --email "sales@acme.com"
```

### Opdrachtgever toevoegen

```
python cli.py clients add "Klant BV" \
    --address "Straat 1, 1000 Brussel" \
    --vat BE987654321 \
    --email "info@klant.be"
```

### Opdrachtgevers importeren uit CSV

```
python cli.py clients import-csv clients.csv \
    --address "Straat 1, 1000 Brussel" \
    --vat BE987654321 \
    --email "info@klant.be"
```

### Opdrachtgevers exporteren naar CSV

```
python cli.py clients export-csv clients_export.csv
```

### Documentnummer prefixen

Bestelbonnen gebruiken doorgaans een nummer dat begint met `BB-`, terwijl
offerteaanvragen `OFF-` gebruiken. De helperfunctie
`_prefix_for_doc_type("Bestelbon")` geeft bijvoorbeeld `BB-` terug zodat deze
prefix automatisch kan worden ingevuld.

### Projectinformatie toevoegen

Gebruik `--project-number` en `--project-name` om projectgegevens op te nemen
in de gegenereerde bestelbonnen of offerteaanvragen:

```
python cli.py copy-per-prod \
    --source src --dest out --bom bom.xlsx --exts pdf \
    --project-number PRJ123 --project-name "Nieuw project"
```

