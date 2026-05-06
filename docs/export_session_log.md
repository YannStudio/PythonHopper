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
- `export`: overzicht van de export zelf, met gegenereerde documenten, statusregels en eventuele padlimiet-waarschuwingen.

Selectiekeys volgen de bestaande keys:

- `production::<naam>`
- `finish::<finish-key>`
- `opticutter::<productie>`

## Laden

Op de bestelbonpagina zijn twee acties voorzien:

- `Exportlog laden`: kies handmatig een JSON-logbestand.
- `Laatste exportlog`: zoekt de nieuwste `Filehopper-exportlog.json` onder de ingestelde exportbestemming.

Als het logbestand `Offerteaanvraag` bevat, vraagt Filehopper of die moeten worden omgezet naar `Bestelbon`. Bij die omzetting worden `OFF-...` documentnummers leeggemaakt, zodat er geen offertenummer als bestelbonnummer meegaat.

Voor het toepassen controleert Filehopper de exportlog tegen de huidige bestelbonpagina:

- BOM-fingerprint gewijzigd.
- Regels uit de exportlog die niet meer bestaan.
- Nieuwe regels in de huidige BOM die geen exportlogwaarden hebben.

Bij verschillen toont Filehopper eerst een controledialoog. De gevonden regels kunnen nog altijd worden toegepast; ontbrekende en nieuwe regels blijven handmatig aanpasbaar.

## Prijzen

De exportlog bewaart prijzen per selectie/bon:

- `unit_price`: eenheidsprijs.
- `total_price`: globale totaalprijs voor die selectie/bon.

Daarnaast kan `items` lijnprijzen per orderregel bevatten. Voor producties en afwerkingen gebruikt Filehopper een stabiele sleutel op basis van PartNumber; voor brutemateriaal op basis van profiel, materiaal en lengte.

Als prijzen ingevuld zijn, voegt de documentgenerator prijsvelden toe aan Excel en PDF. Zonder prijzen blijft de bestaande layout behouden. Lijnprijzen hebben voorrang op de bonbrede eenheidsprijs; een bonbrede totaalprijs blijft als aparte totaalregel zichtbaar.

## Status In De UI

Na het laden van een exportlog krijgen toegepaste regels een `[Log]`-markering op de bestelbonpagina. Bij focus op zo'n rij toont de statusbalk uit welke exportlog de waarden kwamen en of offertes werden omgezet naar bestelbonnen.

## Exportoverzicht

`export.generated_documents` bevat de documenten die Filehopper tijdens de export heeft aangemaakt, met relatieve paden binnen de bundelmap. Orderdocumenten bevatten waar mogelijk ook selectiekey, context, documenttype, documentnummer en leverancier. Daardoor blijft achteraf zichtbaar welke PDF/XLSX-bestanden bij de opgeslagen bestelboninstellingen hoorden.

## Volgende uitbreidingen

- Fijnere merge-opties per veld wanneer een exportlog slechts gedeeltelijk mag worden toegepast.
- Mogelijkheid om vanuit de exportlog snel het vorige exportdocument te openen.
