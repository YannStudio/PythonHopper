"""Help content for the Filehopper settings and quick-start guide."""

QUICK_START_STEPS = [
    {
        "title": "1. Open of laad je data",
        "description": (
            "Begin met het openen van je BOM of andere gegevens in de relevante tab. "
            "Controleer dat klanten, leveranciers en producten correct staan."
        ),
    },
    {
        "title": "2. Controleer je instellingen",
        "description": (
            "Bekijk de instellingen voor bestelbonnen, PDF werkdossiers en exportopties. "
            "Pas indien nodig sjablonen of exportmappen aan."
        ),
    },
    {
        "title": "3. Genereer de documenten",
        "description": (
            "Klik op de export- of genereerknop in de betreffende workflow. "
            "Filehopper maakt nu de bestelbon- of PDF-bestanden voor je."
        ),
    },
    {
        "title": "4. Controleer je output",
        "description": (
            "Open de gegenereerde bestanden in de exportmap en controleer of alles klopt. "
            "Gebruik het logbestand of foutmeldingen als iets niet lukt."
        ),
    },
    {
        "title": "5. Vind extra hulp",
        "description": (
            "Als je vragen hebt of tegen een probleem aanloopt, gebruik dan de FAQ of bekijk de release notes. "
            "Ga voor meer details naar Instellingen > Hulp."
        ),
    },
]

SPARE_PART_QUICK_START_STEPS = [
    {
        "title": "1. BOM voorbereiden",
        "description": (
            "Zet spare-part regels in de BOM op `Production = Spare Parts`. "
            "Gebruik dezelfde propertynamen als in de BOM-template: `Supplier`, "
            "`Supplier code`, `Manufacturer` en `Manufacturer code`."
        ),
    },
    {
        "title": "2. Automatische flow begrijpen",
        "description": (
            "Als je niets aanpast in `Spare parts`, maakt Filehopper zelf een "
            "klaarleglijst, Supplier-groepen, Manufacturer-groepen en eventueel "
            "`Nog toe te wijzen`. Je hoeft dus niet verplicht handmatig groepen te maken."
        ),
    },
    {
        "title": "3. Spare parts controleren",
        "description": (
            "Open het tabblad `Spare parts`. Controleer de status bovenaan en de "
            "groepen links. De kolom `Aandacht` toont hoeveel regels in een groep nog "
            "Supplier/Manufacturer of code-informatie missen. De documentcontrole "
            "toont in een regel of klaarleglijst, groepen, leverancier/fabrikant en "
            "codes klaar zijn."
        ),
    },
    {
        "title": "4. Groepen bijsturen",
        "description": (
            "Selecteer onderdelen rechts en gebruik `Verplaats` om ze in de gekozen "
            "bestelgroep te zetten. Gebruik `Nieuwe groep van selectie` voor een "
            "eigen groep. Met `Selecteer hele groep` kun je een volledige zichtbare "
            "groep in een keer meenemen; met `Selecteer aandacht` pak je alleen de "
            "open of onvolledige regels. Gebruik `Zet op open` voor nog te beslissen "
            "regels en `Terug naar automatische groep` om opnieuw "
            "Supplier/Manufacturer-groepering te gebruiken. Links kun je een groep "
            "ook hernoemen of samenvoegen met een bestaande eigen groep. Dit wijzigt "
            "de originele BOM niet."
        ),
    },
    {
        "title": "5. Aandachtspunten controleren",
        "description": (
            "Klik op `Aandachtspunten` om meteen naar open of onvolledige spare-part "
            "informatie te springen. Controleer vooral regels zonder leverancier, "
            "fabrikant of bestelcode voordat je documenten maakt."
        ),
    },
    {
        "title": "6. Klaarleglijst nakijken",
        "description": (
            "De `Klaarleglijst - alle spare parts` krijgt geen leverancier "
            "in de documentkop. De regels tonen wel `Supplier`, `Supplier code`, "
            "`Manufacturer`, `Manufacturer code`, `Bestelgroep` en `Status`."
        ),
    },
    {
        "title": "7. Documenten klaarmaken",
        "description": (
            "Klik op `Documenten klaarmaken`. Als de documentcontrole nog "
            "aandachtspunten ziet, krijg je eerst een korte bevestiging. Controleer "
            "daarna per spare-parts groep documenttype, leverancier, leveradres, "
            "opmerkingen en prijzen voordat je exporteert."
        ),
    },
    {
        "title": "8. Leverancier wijzigen",
        "description": (
            "Wil je toch bij een andere leverancier bestellen, kies dan in de "
            "bestelbonnenflow een andere leverancier voor de volledige groep. "
            "Voor enkele onderdelen maak je eerst in `Spare parts` een nieuwe "
            "bestelgroep en kies je daarna de gewenste leverancier."
        ),
    },
    {
        "title": "9. Waarschuwingen controleren",
        "description": (
            "Een spare-part `Bestelbon` of `Offerteaanvraag` zonder leverancier "
            "geeft een waarschuwing. Ga je toch door zonder leverancier, dan wordt "
            "die bon overgeslagen. De klaarleglijst mag wel zonder leverancier blijven."
        ),
    },
]

QUICK_MANUAL_SECTIONS = {
    "general": {
        "label": "Algemeen",
        "title": "Quick Start",
        "intro": (
            "Korte gids voor de algemene Filehopper-flow: data laden, instellingen "
            "nakijken en documenten genereren."
        ),
        "steps": QUICK_START_STEPS,
    },
    "spare_parts": {
        "label": "Spare parts",
        "title": "Spare parts quick manual",
        "intro": (
            "Korte gids voor spare parts: automatische groepen begrijpen, "
            "klaarleglijst controleren, leveranciers kiezen en bestel-/offertedocumenten klaarmaken."
        ),
        "steps": SPARE_PART_QUICK_START_STEPS,
    },
}

FAQ_ENTRIES = [
    {
        "question": "Wat doe ik als een exportbestand niet opent?",
        "answer": (
            "Controleer eerst of het bestand daadwerkelijk is aangemaakt in de exportmap. "
            "Open daarna het bestand handmatig vanuit Verkenner. Als het bestand niet bestaat, controleer dan "
            "of de export vanuit Filehopper zonder fouten is afgerond en of de juiste exportmap is ingesteld."
        ),
    },
    {
        "question": "Waarom wordt mijn leverancier niet gevonden?",
        "answer": (
            "Controleer of de leveranciersnaam exact overeenkomt met de naam in de leverancierendatabase. "
            "Soms helpt het om hoofdletters, leestekens of spaties te controleren. Indien nodig kun je "
            "de leverancier opnieuw toevoegen onder Leveranciersbeheer."
        ),
    },
    {
        "question": "Wat als de release notes niet geladen worden?",
        "answer": (
            "Filehopper zoekt naar een lokale CHANGELOG.md in de projectmap. Als die ontbreekt of niet toegankelijk is, "
            "dan zie je een foutmelding. Controleer of het bestand aanwezig is of gebruik de app vanuit de juiste projectroot."
        ),
    },
    {
        "question": "Hoe maak ik een back-up van mijn gegevens?",
        "answer": (
            "Maak een kopie van je runtime JSON-bestanden in de map waar Filehopper draait, zoals clients_db.json, "
            "suppliers_db.json en app_settings.json. Bewaar deze kopieën op een veilige plaats voordat je wijzigingen aanbrengt."
        ),
    },
    {
        "question": "Ik krijg een foutmelding bij het genereren van een PDF werkdossier. Wat nu?",
        "answer": (
            "Controleer eerst de invoergegevens en exportinstellingen. Als het probleem blijft bestaan, kijk dan in het debuglogbestand "
            "voor details en controleer of alle benodigde velden zijn ingevuld."
        ),
    },
]
