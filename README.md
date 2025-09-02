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

## Custom BOM

Het GUI-programma bevat een tabblad **Custom BOM** waarmee je snel een
eigen stuklijst kunt opstellen. Start de applicatie zonder CLI-argumenten
(`python main.py`) en open het tabblad *Custom BOM*.

1. **Data invoeren of plakken** – de tabel bevat de kolommen
   `PartNumber`, `Description`, `Material`, `Quantity` en `Production`.
   Kopieer cellen uit een spreadsheet en plak ze rechtstreeks met
   `Ctrl+V`/`Cmd+V`; elke kolom wordt automatisch toegewezen.
2. **Gebruik BOM** – kies *Gebruik BOM* om de ingevoerde gegevens naar het
   hoofdscherm te exporteren en de stuklijst verder te verwerken.

Deze functie maakt gebruik van de `tksheet`-bibliotheek (opgenomen in
`requirements.txt`). Zorg ervoor dat deze dependency is geïnstalleerd en
hou er rekening mee dat alleen de bovenstaande kolommen worden ondersteund.

