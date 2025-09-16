# Filehopper
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

