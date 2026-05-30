# Release Process - Auto-Generated Release Notes

## Overview

Het release proces is nu geautomatiseerd met auto-generated release notes uit Git commits. Dit uses **Conventional Commits** format.

## Conventional Commits Format

Commits worden geparse op basis van dit format:

```
type(scope): description

optionele body

optionele footer
```

### Type (Verplicht)
- `feat` - Nieuwe feature
- `fix` - Bug fix
- `docs` - Documentatie
- `style` - Code style (geen logische wijzigingen)
- `refactor` - Code refactoring
- `perf` - Performance improvements
- `test` - Tests toevoegen/aanpassen
- `chore` - Build, dependencies, etc.
- `ci` - CI/CD configuratie

### Scope (Optioneel)
Component of module die wordt aangepast, bijv:
- `(orders)` - orders module
- `(gui)` - GUI module
- `(bom)` - BOM handling

### Description (Verplicht)
Korte beschrijving (imperatief, geen hoofdletter, geen punt).

### Breaking Changes
Markeer met `!` na scope of gebruik `BREAKING CHANGE:` in footer:

```
feat(api)!: change authentication method

BREAKING CHANGE: old authentication no longer supported
```

## Voorbeelden van Conventional Commits

```
feat(manual-order): add template support
fix(pdf): correct page break positioning
refactor(orders): extract color utilities
docs: update installation instructions
test(bom): verify sync functionality
chore(deps): update pandas to 2.0
ci: add GitHub Actions workflow
```

## Usage

### 1. Preview Release Notes

Voordat je een release doet, bekijk wat wordt gegenereerd:

```bash
python scripts/preview_release_notes.py 3.3
```

Output:
```
📋 Release notes preview for 3.3
📅 Generated from commits since: v3.2

### ✨ Features
- enhance manual order tab templates (15f9e1a)
- prepend related BOM pdfs to exports (6f14c26)

### 🐛 Bug Fixes
- correct page break positioning (abc1234)

### ♻️ Refactoring
- extract CLI to dedicated module (689cdb6)

✓ Found 12 commits to include
```

### 2. Prepare Release

Standaard release voorbereiding (met auto-generated notes):

```bash
python scripts/release.py 3.3
```

Dit zal:
1. ✓ APP_VERSION in app_paths.py updaten
2. ✓ Windows version files updaten
3. ✓ CHANGELOG.md updaten met auto-generated release notes

### 3. Preview Notes Before Updating

Release notes controleren voordat ze in CHANGELOG worden geschreven:

```bash
python scripts/release.py 3.3 --preview-notes
```

Dit toont eerst een preview, dan vraagt het om bevestiging.

### 4. Met Tests

```bash
python scripts/release.py 3.3 --test
```

Voert tests uit na versie-update.

### 5. Met Build

```bash
python scripts/release.py 3.3 --build --target windows
python scripts/release.py 3.3 --build --target macos --onefile
```

Build executables na versie-update.

### 6. Alles Combineren

```bash
python scripts/release.py 3.3 --preview-notes --test --build --target windows --onefile
```

## Voorbeeld: Complete Release Workflow

```bash
# 1. Preview wat komt
python scripts/preview_release_notes.py 3.3

# 2. Ziet er goed uit, update alles
python scripts/release.py 3.3 --test --build --target windows --onefile

# 3. Review changes
git diff

# 4. Commit en tag
git commit -am "Release 3.3"
git tag v3.3

# 5. Push
git push && git push --tags
```

## Generated CHANGELOG.md Format

Changelog entries worden automatisch gegenereerd:

```markdown
## 3.3 - 2026-05-30

### ✨ Features
- enhance manual order tab templates (15f9e1a)
- prepend related BOM pdfs to combined exports (6f14c26)

### 🐛 Bug Fixes
- correct page break positioning (abc1234)

### ♻️ Refactoring
- extract CLI to dedicated module (689cdb6)

## 3.2 - 2026-05-29
...
```

## Best Practices

### ✓ Do's
- Schrijf commits met Conventional Commits format
- Gebruik imperatief mood: "add feature", niet "added feature"
- Maak commits klein en gefocust
- Beschrijf WHY in de commit body, niet WHAT (dat is in de code)
- Markeer breaking changes expliciet

### ✗ Don'ts
- Geen willekeurige commit messages: "update stuff", "fixes", "asdf"
- Geen monstercommits met 500 files
- Geen typo's in commit messages (moeilijk te fixen later)

## Troubleshooting

### Git staat niet in PATH
```bash
# Zorg dat git beschikbaar is
git --version
```

### Geen tags in repository
```bash
# Preview_release_notes zal dan alle commits pakken
# Dit is OK voor de eerste release
```

### Release notes zien er raar uit
```bash
# Check commit messages
git log --oneline

# Format moet zijn: type(scope): description
```

## Files Gerelateerd aan Release

- `scripts/release.py` - Hoofd release script
- `scripts/release_notes_generator.py` - Auto-generate release notes uit commits
- `scripts/preview_release_notes.py` - Preview tool
- `CHANGELOG.md` - Gegenereerde changelog
- `app_paths.py` - Bevat APP_VERSION

## Voordelen

✅ **Consistent format** - Alle release notes ziet er hetzelfde uit  
✅ **Geen handmatig werk** - Notes worden geautomatiseerd gegenereerd  
✅ **Beter tracking** - Commits worden gecategoriseerd  
✅ **Audit trail** - Commit hashes in release notes  
✅ **Breaking changes** - Duidelijk gemarkeerd  
✅ **Makkelijk om terug te zien** - Changelog is altijd up-to-date  

## Volgende Stappen

1. Zorg dat je commits Conventional Commits format volgen
2. Test met `python scripts/preview_release_notes.py 3.3`
3. Gebruik `python scripts/release.py 3.3` voor releases
4. Git tag en push!
