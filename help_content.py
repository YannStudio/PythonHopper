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
