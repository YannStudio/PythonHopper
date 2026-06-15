# Changelog

## 3.3 - 2026-06-05

- Spare-parts-regels in de BOM krijgen een eigen flow met volledige monteurslijst, groepering per leverancier/fabrikant en bestelgerichte codes op bonnen en exports.
- Spare-parts-groepen kunnen nu met beheerbare presets voorbereid worden; exportlogs bewaren en herstellen ook de spare-part-verdeling.
- Spare-part bestelbonnen/offerteaanvragen zonder leverancier geven nu een bevestiging vlak voor export.
- Exportlog-controle toont herstelbare custom spare-partgroepen nu apart van echt ontbrekende bestelbonregels.
- Spare-partsverdelingen uit exportlogs worden robuuster hersteld wanneer BOM-rijindexen verschuiven.
- Spare-part identity keys bewaren lege veldposities, zodat toekomstig exportlog-herstel minder ambigu is.
- Handmatige en preset spare-partgroepen blijven beter behouden bij het verversen van een BOM met verschoven rijindexen.
- De volledige spare-partlijst wordt in UI/export duidelijker als klaarleglijst benoemd.
- De spare-part klaarleglijst gedraagt zich in de Bestelbonnen-flow niet meer als leveranciersbon.
- De spare-part klaarleglijst kan niet meer als gekoppelde bon of groepsmaster gebruikt worden.
- De exportlaag forceert de spare-part klaarleglijst nu altijd als standaarddocument zonder leverancier.
- Het PDF werkdossier heeft nu een aparte keuze om de volledige spare-parts klaarleglijst mee in te voegen.
- Het PDF werkdossier meldt nu ook wanneer de aangevinkte spare-parts klaarleglijst nog niet voorbereid is.
- Diagnostiek waarschuwt nu ook voor lege of overlappende spare-part presetregels.
- PDF dossier wist na het kiezen van een volgordepreset de gearceerde comboboxselectie, zoals bij de moduskeuze.
- Informatietekst bij de PDF dossier-modus en volgordepreset gebruikt nu dezelfde tekstkleur.
- Bestelbon- en werkdossieruitvoer verwerkt maatvelden, totalen en plaat-/profielgegevens consequenter in Excel en PDF.
- Lange bestelbontabellen en celteksten breken beter af, zodat totalen en details leesbaar blijven.
- Leveranciersbeheer bewaart extra adres- en contactgegevens correct bij het bewerken.
- Leveranciersdata en templategegevens zijn bijgewerkt.
- GitHub Actions workflow toegevoegd voor automatische Python-tests.

## 3.2 - 2026-06-01

- PDF werkdossier werkt nu in dezelfde tab in plaats van via een pop-up.
- PDF werkdossier heeft aparte subtabs voor dossier maken en presetregels beheren.
- Presetregels gebruiken nu een algemene knop om productieblokken toe te voegen, met duidelijkere benaming voor overige producties.
- PDF dossier toont vooraf de mergevolgorde met categorie, productie, documenttype en bestandsnaam.
- PDF dossier toont de exportnaam en exportmap voordat de combineeractie start.
- PDF combineren blokkeert dubbele klikken tijdens een lopende actie en toont voortgang in procenten.
- Bestelbonnen voor PDF dossier hebben een aparte voorbereidingsflow: de app opent de Bestelbonnen-tab in PDF-dossiercontext, maakt bon-PDF's in een zichtbare projectmap en keert automatisch terug naar PDF werkdossier.
- Bonnen zonder tekening, zoals afwerkingsbonnen of stock/sparepart-bonnen, worden achteraan aan het PDF dossier toegevoegd onder "Aanvullende bonnen zonder tekening".
- Leveranciers zoeken en selecteren in de Bestelbon-editor is hersteld; dropdownselecties worden niet meer afgebroken door focusverlies.
- Preset editor laadt presetwaarden direct bij het openen van de Presetregels-tab zonder dat je hoeft op de dropdown te klikken.
- Blanco template blijft nu geselecteerd in plaats van terug te springen naar "Werkdossier standaard".

## 3.1 - 2026-05-21

- Filehopper 3.0 branch gemerged naar main.
- Releaseversie verhoogd naar 3.1.
