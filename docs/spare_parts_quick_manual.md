# Spare parts quick manual

Deze korte handleiding ondersteunt de gebruiker bij de spare-parts module in Filehopper.

## Doel

Gebruik `Spare parts` voor BOM-regels die apart klaargelegd, besteld of opgevraagd moeten worden. Een regel hoort in deze flow wanneer `Production` gelijk is aan `Spare Parts`.

## BOM-kolommen

Houd de BOM-propertynamen overal hetzelfde. De spare-parts tab, presets en exportdocumenten gebruiken deze namen:

- `PartNumber`
- `Description`
- `Aantal`
- `Supplier`
- `Supplier code`
- `Manufacturer`
- `Manufacturer code`

Filehopper kan enkele oude Nederlandstalige aliassen nog inlezen, maar nieuwe BOMs en templates gebruiken best de propertynamen hierboven.

## Dagelijkse flow

1. Laad de BOM in Filehopper.
2. Open het tabblad `Spare parts`.
3. Controleer de statusregel bovenaan. Los vooral regels in `Nog toe te wijzen`, `Mist Supplier/Manufacturer` en `Mist Supplier code/Manufacturer code` op.
4. Controleer links de groepen. Filehopper maakt automatisch:
   - `Volledige lijst` voor de klaarleglijst.
   - Supplier-groepen voor regels met `Supplier`.
   - Manufacturer-groepen als fallback wanneer `Supplier` leeg is.
5. Selecteer onderdelen rechts en kies bovenaan een `Bestelgroep` als je handmatig wilt hergroeperen.
6. Gebruik `Auto` om geselecteerde regels terug automatisch te laten groeperen.
7. Gebruik `Presets toepassen` voor terugkerende regels.
8. Klik `Documenten klaarmaken`.
9. Controleer per spare-parts groep documenttype, leverancier, leveradres, opmerkingen en prijzen.
10. Exporteer de documenten.

## Presets

Presets groeperen regels automatisch op basis van een BOM-property. Gebruik bij voorkeur deze matchvelden:

- `Supplier`
- `Supplier code`
- `Manufacturer`
- `Manufacturer code`

Voorbeelden:

- `Supplier` bevat `RS` -> doelgroep `Electro`
- `Manufacturer` is `Festo` -> doelgroep `Pneumatica`
- `Manufacturer code` begint met `SM-` -> doelgroep `Mechanisch`

## Belangrijke afspraken

- Verander `Production` niet naar een bestelgroep. Laat die waarde `Spare Parts`.
- Gebruik `Supplier` en `Supplier code` als er rechtstreeks bij een leverancier besteld wordt.
- Gebruik `Manufacturer` en `Manufacturer code` voor fabrikantinformatie, ook als de bestelling via een leverancier loopt.
- De volledige lijst is een klaarleglijst en krijgt standaard geen leverancier.
- Groepskeuzes wijzigen de originele BOM niet; ze worden in Filehopper en de exportlog bewaard.
