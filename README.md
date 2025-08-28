# PythonHopper
Filehopper in Python

## Voorbeelden

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

### Bestelbon, offerte of offerteaanvraag genereren

Met het commando `copy-per-prod` worden bestanden per productie gekopieerd en wordt automatisch een document aangemaakt. Via de optie `--doc-type` bepaal je of er een **bestelbon** (standaard), een **offerte** of een **offerteaanvraag** gemaakt wordt. Bij een offerteaanvraag kun je ook een antwoorddeadline meegeven; deze verschijnt in de gegenereerde PDF en Excel.

```
python cli.py copy-per-prod --source bron --dest doel --bom bom.csv --exts pdf \
    --doc-type offerteaanvraag
```

