# Export Session Log

Filehopper schrijft bij een succesvolle bestelbonexport een projectgebonden logbestand:

`Filehopper-exportlog.json`

Dit bestand staat in de exportbundelmap en bewaart de bestelbonkeuzes die normaal alleen in de GUI-state zitten.

## Doel

- Een offerte-export later opnieuw inladen.
- Dezelfde leveranciers, documenttypes, nummers, leveradressen, opmerkingen, exportvlaggen, EN 1090-keuzes en prijzen opnieuw toepassen.
- Offerteaanvragen kunnen omzetten naar effectieve bestelbonnen zonder de selectie opnieuw op te bouwen.

## Inhoud

Het logbestand bevat:

- `schema_version`: versie van het logformaat.
- `app`: Filehopper-versie.
- `project`: projectnummer, projectnaam en opdrachtgever.
- `bom`: bronpad, bestandsnaam, rijtelling en fingerprint van relevante BOM-data.
- `order_state`: de bestelbonstate per selectiekey.

Selectiekeys volgen de bestaande keys:

- `production::<naam>`
- `finish::<finish-key>`
- `opticutter::<productie>`

## Laden

Op de bestelbonpagina zijn twee acties voorzien:

- `Exportlog laden`: kies handmatig een JSON-logbestand.
- `Laatste exportlog`: zoekt de nieuwste `Filehopper-exportlog.json` onder de ingestelde exportbestemming.

Als het logbestand `Offerteaanvraag` bevat, vraagt Filehopper of die moeten worden omgezet naar `Bestelbon`. Bij die omzetting worden `OFF-...` documentnummers leeggemaakt, zodat er geen offertenummer als bestelbonnummer meegaat.

## Prijzen

De eerste implementatie bewaart prijzen per selectie/bon:

- `unit_price`: eenheidsprijs.
- `total_price`: globale totaalprijs voor die selectie/bon.

Als prijzen ingevuld zijn, voegt de documentgenerator prijsvelden toe aan Excel en PDF. Zonder prijzen blijft de bestaande layout behouden.

## Volgende uitbreidingen

- Lijnprijzen per PartNumber of profielregel.
- Validatiedialoog bij BOM-wijzigingen met overzicht van niet-gematchte en nieuwe regels.
- Duidelijke visuele status op de bestelbonpagina wanneer een regel uit een exportlog komt.
