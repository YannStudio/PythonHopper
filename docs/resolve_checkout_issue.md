# Checkout-probleem met `app_settings.json`

Wanneer je tijdens het wisselen van branch de melding

```
app_settings.json: needs merge
error: you need to resolve your current index first
```

krijgt, betekent dit dat Git een onafgehandelde merge-conflict in `app_settings.json` ziet. Dit gebeurt vaak als je eerder een merge startte of bestanden wijzigde zonder de merge af te ronden.

## Oplossing

1. **Controleer de status**

   ```powershell
   git status
   ```

   Als `app_settings.json` nog in de lijst met conflicten staat, open het bestand dan en verwijder eventuele conflictmarkeringen (`<<<<<<<`, `=======`, `>>>>>>>`).

2. **Bewaar de gewenste inhoud**

   Kies welke versie van de instellingen je wil behouden en pas het JSON-bestand daarop aan. Zorg dat het geldige JSON blijft.

3. **Markeer het conflict als opgelost**

   ```powershell
   git add app_settings.json
   ```

4. **Maak desnoods een commit**

   ```powershell
   git commit -m "Los merge-conflict in app_settings.json op"
   ```

   Hiermee rond je de merge af. Als je geen merge bezig was maar enkel lokale wijzigingen wilde weggooien, kun je ook eerst terug naar de laatst bekende versie gaan:

   ```powershell
   git checkout -- app_settings.json
   ```

5. **Schakel naar de gewenste branch**

   ```powershell
   git checkout main
   ```

## Opmerking over branch-namen

Je gaf aan dat je per ongeluk aan `Main` werkte en de branch later `fout` hebt genoemd. Controleer of `main` daadwerkelijk bestaat (bijv. `git branch -a`). Als `main` niet bestaat, kun je hem opnieuw aanmaken:

```powershell
git branch main origin/main
```

Pas daarna kun je `git checkout main` uitvoeren.

Met bovenstaande stappen zou het wisselen van branch weer moeten lukken.
