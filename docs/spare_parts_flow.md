# Spare parts flow

Deze flow is bedoeld voor BOM-regels die niet bij een productieblok horen, maar als spare parts moeten worden klaargelegd, besteld of opgevraagd.

## BOM-gegevens

Een regel wordt als spare part gezien wanneer de kolom `Production` een waarde zoals `Spare Parts`, `Spare Part` of `SpareParts` bevat.

Ondersteunde kolommen:

- `PartNumber`: artikelnummer uit de BOM.
- `Description`: omschrijving voor de bon en de montagelijst.
- `Aantal`: aantal stuks.
- `Materiaal`: optionele materiaalinfo.
- `Supplier`: voorkeursleverancier.
- `Supplier code`: artikelcode bij de leverancier.
- `Manufacturer`: fabrikant wanneer er geen duidelijke leverancier is.
- `Manufacturer code`: fabrikantcode.

## Groepering

Filehopper maakt automatisch meerdere spare-parts-selecties:

- `Volledige lijst`: alle spare parts samen, standaard als `Standaard bon`. Dit is de lijst voor monteurs om klaar te leggen voor vertrek.
- Leveranciersgroepen: alle spare parts met dezelfde `Supplier`, standaard als `Bestelbon`.
- Fabrikantgroepen: fallback wanneer `Supplier` leeg is maar `Manufacturer` wel ingevuld is.
- `Zonder leverancier/fabrikant`: controlelijst voor regels die nog onvoldoende bestelinfo hebben.

Leveranciersgroepen krijgen voorrang op fabrikantgroepen. Dat houdt de bestelstroom praktisch: wat rechtstreeks bij een leverancier besteld wordt, komt niet dubbel in een fabrikantgroep terecht.

## GUI-flow

1. Laad of bewerk de BOM zoals gewoonlijk.
2. Open het tabblad `Spare parts` om de volledige lijst en de groepen te controleren.
3. Selecteer een of meerdere regels en zet ze met `Zet selectie` in een andere bestelgroep.
4. Maak met `Nieuwe groep` een vrije groep, bijvoorbeeld `Electro`.
5. Gebruik `Terug open` om regels naar `Nog toe te wijzen` te zetten.
6. Gebruik `Auto` om geselecteerde regels terug volgens Supplier/Manufacturer te laten groeperen.
7. Gebruik `Bestelbonnen klaarmaken` om naar dezelfde bestelbonnenflow te gaan als voor producties, finishes en brutemateriaal.
8. Kies per spare-parts-groep of het document geexporteerd moet worden.
9. Kies per groep het documenttype, documentnummer, leverancier, leveradres, opmerkingen, prijzen en BTW.
10. Start de export. De volledige lijst en de gekozen bestel-/offertedocumenten worden samen met de gewone export aangemaakt.

Handmatige groepskeuzes blijven in de app-state zolang de huidige BOM actief is. Ze wijzigen de originele BOM-kolom `Production` niet.

## Documenten en export

Spare-parts-documenten gebruiken een eigen tabelindeling met bestelgerichte kolommen:

- Artikel nr.
- Omschrijving
- St.
- Supplier
- Supplier code
- Fabrikant
- Fabrikant code
- Bestelgroep
- Status

Prijsvelden worden alleen toegevoegd wanneer er prijsgegevens ingevuld zijn. De exportlog bewaart de keuzes met `sparepart::...` selectiekeys, zodat dezelfde instellingen later opnieuw toegepast kunnen worden.

## Praktische afspraken

- Laat `Production` op `Spare Parts` staan; verander deze waarde niet meer handmatig naar `Electro`, `Herbaroof` of een andere bestelgroep.
- Vul bij voorkeur `Supplier` en `Supplier code` in wanneer je effectief bij een leverancier bestelt.
- Gebruik `Manufacturer` en `Manufacturer code` wanneer de fabrikantcode belangrijk is, ook als de bestelling via een leverancier loopt.
- Controleer de groep met ontbrekende data voor je de finale export maakt.
