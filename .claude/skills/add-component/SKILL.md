---
name: add-component
description: Scaffold a new React component following project patterns
disable-model-invocation: true
allowed-tools: Read, Write, Edit, Grep, Glob
---

# Add Component

Create a new React component named `$ARGUMENTS` following project patterns.

## Steps

1. Determine if this should be a Server or Client component based on the name and likely usage.

2. Check existing components in `web/src/components/` for patterns:
   - How imports are structured
   - How props are typed
   - How Tailwind classes are applied
   - How `cn()` is used for conditional styling

3. Create `web/src/components/$ARGUMENTS.tsx` with:
   - TypeScript interface for props
   - Named export (not default)
   - `"use client"` directive only if the component needs interactivity
   - Tailwind CSS for styling using `cn()` for conditionals
   - Proper React 19 patterns (no forwardRef, use `use()` for context)

4. If the component needs types, add them to `web/src/types/index.ts`

5. Report what was created and suggest where to use it.
