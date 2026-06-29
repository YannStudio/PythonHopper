# Spare parts quick manual

Deze handleiding beschrijft de aanbevolen flow voor spare parts in Filehopper:
eerst de spare-parts groepen controleren, daarna leveranciers en documenten
kiezen, en pas op het einde exporteren.

## Doel

Gebruik `Spare parts` voor BOM-regels die apart klaargelegd, besteld of
opgevraagd moeten worden. Een regel hoort in deze flow wanneer `Production`
gelijk is aan `Spare Parts`.

## BOM-kolommen

Houd de BOM-propertynamen overal hetzelfde. De spare-parts tab en
exportdocumenten gebruiken deze namen:

- `PartNumber`
- `Description`
- `Aantal`
- `Supplier`
- `Supplier code`
- `Manufacturer`
- `Manufacturer code`

Filehopper kan enkele oude Nederlandstalige aliassen nog inlezen, maar nieuwe
BOMs en templates gebruiken best de propertynamen hierboven.

## Belangrijk uitgangspunt

Als je niets aanpast in het tabblad `Spare parts`, blijft de flow automatisch
werken. Filehopper maakt dan zelf groepen op basis van de ingevulde BOM-data:

- `Volledige lijst`: de klaarleglijst voor alle spare parts samen.
- Supplier-groepen: regels met dezelfde `Supplier`.
- Manufacturer-groepen: fallback wanneer `Supplier` leeg is maar
  `Manufacturer` wel ingevuld is.
- `Nog toe te wijzen`: regels met te weinig bestelinfo.

Je moet dus niet verplicht handmatig groepen maken. De spare-parts flow is wel
een voorbereidingsstap voor de bestelbonnen: net zoals Opticutter eerst zijn
zaagkeuzes nodig heeft, controleer je hier eerst de spare-parts groepen voordat
je de documenten definitief exporteert.

## Aanbevolen stappen

1. Laad de BOM in Filehopper.
2. Open het tabblad `Spare parts`.
3. Controleer de statusregel bovenaan.
4. Controleer links de groepen en de kolom `Mist`.
5. Selecteer links een groep om rechts de onderdelen in die groep te bekijken.
6. Laat de automatische groepering staan als de groepen al kloppen.
7. Selecteer onderdelen rechts als je ze toch anders wilt indelen.
8. Kies of typ bovenaan een `Bestelgroep` en klik `Zet selectie`.
9. Gebruik `Nieuwe groep` om snel een eigen groep te maken.
10. Gebruik `Auto` om geselecteerde regels terug naar automatische
    Supplier/Manufacturer-groepering te zetten.
11. Gebruik `Selecteer aandacht`, `Groep hernoemen` en `Samenvoegen met...`
    om uitzonderingen snel af te werken.
12. Klik `Documenten klaarmaken`.
13. Controleer in de bestelbonnenflow per spare-parts groep:
    documenttype, leverancier, documentnummer, leveradres, opmerkingen,
    prijzen en BTW.
14. Exporteer de documenten.

## Klaarleglijst

De `Volledige lijst` is een klaarleglijst voor de monteurs. Die lijst is geen
leveranciersbon en krijgt daarom geen leverancier in de documentkop.

De regels op de klaarleglijst tonen wel de bestelinfo die nodig is om onderdelen
terug te vinden of leveringen te controleren:

- `Supplier`
- `Supplier code`
- `Manufacturer`
- `Manufacturer code`
- `Bestelgroep`
- `Status`

Zo kan de klaarleglijst gebruikt worden om onderdelen fysiek klaar te leggen,
maar ook om tussen leveringen of fabrikantcodes te zoeken.

## Leveranciers kiezen

Leveranciers kies je pas in de bestelbonnenflow, na `Documenten klaarmaken`.

- Supplier-groepen krijgen normaal automatisch de leverancier uit de BOM.
- Manufacturer-groepen hebben vaak nog geen echte leverancier en moeten daar
  manueel een leverancier krijgen.
- Custom groepen krijgen geen vaste leverancier uit de BOM en moeten meestal
  manueel gekozen worden.
- De klaarleglijst blijft `Standaard bon` en heeft geen leverancier nodig.

Wil je onderdelen die in de BOM bij leverancier A staan toch bij leverancier B
bestellen, dan kan dat op twee manieren:

1. Voor de volledige groep kies je bij `Documenten klaarmaken` gewoon een
   andere leverancier.
2. Voor enkele regels selecteer je die regels in `Spare parts`, zet je ze in
   een nieuwe bestelgroep, en kies je daarna in de bestelbonnenflow de gewenste
   leverancier voor die groep.

Deze keuzes wijzigen de originele BOM niet. Ze worden alleen in Filehopper en
in de exportlog bewaard.

## Waarschuwingen en meldingen

De spare-parts flow probeert adviserend te waarschuwen zonder de gebruiker
onnodig te blokkeren.

- De statusregel bovenaan meldt hoeveel spare parts en groepen er zijn.
- De kolom `Mist` toont hoeveel regels in een groep nog info missen.
- `Nog toe te wijzen` wijst op regels zonder duidelijke Supplier/Manufacturer.
- `Mist Supplier/Manufacturer` betekent dat er geen route-informatie is.
- `Mist Supplier code/Manufacturer code` betekent dat er wel een route is, maar
  nog geen artikelcode.
- Een spare-part `Bestelbon` of `Offerteaanvraag` zonder leverancier geeft een
  waarschuwing voor export.
- Als je toch doorgaat zonder leverancier, wordt die leveranciersbon
  overgeslagen met een statusmelding.
- De klaarleglijst mag zonder leverancier blijven.

## Groepen beheren

Filehopper maakt automatisch groepen op basis van `Supplier` en daarna
`Manufacturer`. Voor projectgebonden uitzonderingen gebruik je de handmatige
groepsacties:

- `Verplaats` zet geselecteerde regels in de gekozen bestelgroep.
- `Nieuwe groep van selectie` maakt meteen een eigen groep.
- `Selecteer hele groep` en `Selecteer aandacht` versnellen bulkwerk.
- `Groep hernoemen` en `Samenvoegen met...` ruimen bestaande eigen groepen op.
- `Terug naar automatische groep` verwijdert de handmatige keuze voor de
  geselecteerde regels.

## Belangrijke afspraken

- Verander `Production` niet naar een bestelgroep. Laat die waarde
  `Spare Parts`.
- Gebruik `Supplier` en `Supplier code` als er rechtstreeks bij een leverancier
  besteld wordt.
- Gebruik `Manufacturer` en `Manufacturer code` voor fabrikantinformatie, ook
  als de bestelling via een leverancier loopt.
- Maak bestelbonnen pas nadat de spare-parts groepen juist staan.
- De volledige lijst is een klaarleglijst en krijgt standaard geen leverancier.
- Groepskeuzes wijzigen de originele BOM niet; ze worden in Filehopper en de
  exportlog bewaard.
