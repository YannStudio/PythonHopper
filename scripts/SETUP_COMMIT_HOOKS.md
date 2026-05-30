# Setup Commit Hooks for Release Notes Generation

Dit document beschrijft hoe je lokale commit hooks inricht zodat commit messages worden gevalideerd tegen Conventional Commits format.

## Quick Setup (Git 2.9+)

### 1. Configure Git hooks directory

```bash
git config core.hooksPath .githooks
```

Dit vertelt Git om hooks uit `.githooks/` directory te gebruiken.

### 2. Make hook executable

```bash
chmod +x .githooks/prepare-commit-msg
```

### 3. Test it

Probeer een commit met verkeerd format:

```bash
git commit -m "fixed stuff"
```

Je zal een foutmelding krijgen:
```
❌ Commit message validation failed:
Invalid commit message format. Expected: type(scope): description
  Got: fixed stuff
  Types: feat, fix, docs, style, refactor, perf, test, chore, ci, revert
  Scope: optional, in kebab-case
  Example: feat(orders): add new export format
```

Probeer nu met correct format:

```bash
git commit -m "fix(orders): resolve pdf export bug"
```

Dit zal werken!

## Commit Message Examples

### ✅ Valid Examples

```bash
# Simple feature
git commit -m "feat: add dark mode support"

# Feature with scope
git commit -m "feat(gui): add dark mode support"

# Bug fix
git commit -m "fix(pdf): correct page break calculation"

# Documentation
git commit -m "docs: update installation guide"

# Refactoring
git commit -m "refactor(orders): extract color utilities"

# Tests
git commit -m "test(bom): verify sync functionality"

# With body explaining why
git commit -m "feat(export): support new file format

The legacy format is no longer used by customers,
so we're adding support for the new industry standard."

# Breaking changes
git commit -m "feat(api)!: change authentication method

BREAKING CHANGE: Legacy tokens are no longer supported.
Use new OAuth2 flow instead."
```

### ❌ Invalid Examples

```bash
# No type
git commit -m "update stuff"

# Wrong type
git commit -m "bugfix: fix something"

# Description capitalized
git commit -m "feat: Add Feature"

# Description ends with period
git commit -m "feat: add feature."

# Scope wrong format (should be kebab-case)
git commit -m "feat(myScope): add feature"
```

## Disable Hooks Temporarily

Als je een commit moet doen zonder validatie (niet aanbevolen):

```bash
git commit --no-verify -m "your message"
```

## Manual Validation

Je kan ook handmatig commits valideren:

```bash
python scripts/preview_release_notes.py 3.3
```

## Troubleshooting

### Hook niet wordt uitgevoerd

1. Check dat `.githooks` directory bestaat
2. Zorg dat hook executable is: `chmod +x .githooks/prepare-commit-msg`
3. Zorg dat Git hooks path geconfigureerd is:
   ```bash
   git config core.hooksPath .githooks
   ```

### "Permission denied" error

```bash
chmod +x .githooks/prepare-commit-msg
```

### Merge commits geven error

Merge commits worden genegeerd en hoeven niet aan format te voldoen.

## IDE Integration

### VS Code

Installeer de **Conventional Commits** extension:
- ID: `vivaxy.vscode-conventional-commits`

Dit geeft je een UI om commits met correct format in te voeren.

### JetBrains IDEs (PyCharm, etc.)

Installeer **Conventional Commit** plugin:
- Ga naar: Settings → Plugins → Browse Repositories
- Zoek: "Conventional Commit"
- Install en restart IDE

## CI/CD Integration

Als je CI/CD hebt, kan je ook server-side validatie doen met commitlint:

```bash
# Install commitlint and husky
npm install commitlint @commitlint/config-conventional husky --save-dev

# Setup husky
npx husky install

# Test
git commit -m "test"  # Will be validated
```

## Next Steps

1. Run git config command hierboven
2. Make hook executable
3. Test met `git commit -m "feat: test commit"`
4. Start using Conventional Commits!

## Resources

- [Conventional Commits](https://www.conventionalcommits.org/)
- [commitlint](https://commitlint.js.org/)
- [Semantic Versioning](https://semver.org/)
