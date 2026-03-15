---
name: lint
description: Run ESLint on the web app and auto-fix issues
disable-model-invocation: true
allowed-tools: Bash(npm run lint), Bash(npx eslint *), Read, Edit
---

# Lint

Run ESLint on the web frontend and fix any issues.

## Steps

1. Run ESLint:
   ```bash
   cd web && npm run lint
   ```

2. If there are fixable errors, attempt auto-fix:
   ```bash
   cd web && npx eslint --fix src/
   ```

3. For remaining errors that can't be auto-fixed:
   - Read the failing files
   - Apply manual fixes following the project's code style rules
   - Re-run lint to verify

4. Report what was fixed
