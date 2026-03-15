---
name: fix-issue
description: Fix a GitHub issue by number
disable-model-invocation: true
allowed-tools: Bash(gh *), Read, Edit, Write, Grep, Glob, Bash(npm run lint), Bash(npm run build), Bash(git *)
---

# Fix GitHub Issue

Fix GitHub issue #$ARGUMENTS.

## Steps

1. Fetch the issue details:
   ```bash
   gh issue view $ARGUMENTS
   ```

2. Understand the requirements from the issue description and comments.

3. Explore the relevant code to understand the current implementation.

4. Implement the fix:
   - Make targeted, minimal changes
   - Follow the project's code conventions
   - Add types where needed

5. Verify the fix:
   ```bash
   cd web && npm run lint && npm run build
   ```

6. Create a commit with a descriptive message referencing the issue:
   ```
   fix: <description> (closes #$ARGUMENTS)
   ```
