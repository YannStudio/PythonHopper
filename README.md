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

