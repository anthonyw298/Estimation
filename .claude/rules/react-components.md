---
paths:
  - "web/src/components/**/*.tsx"
  - "web/src/app/**/*.tsx"
---

# React Component Rules

## Server vs Client Components
- Default to Server Components (no directive needed)
- Add `"use client"` only for: event handlers, useState/useEffect, browser APIs
- Keep client components small -- push logic up to server components

## React 19 Patterns
- Use `use()` hook for reading promises and context (replaces useContext)
- No `forwardRef` -- ref is a regular prop in React 19
- Use `useActionState` for form state management
- Use `useOptimistic` for optimistic updates
- Prefer Server Actions (`"use server"`) for mutations

## Component Structure
```tsx
// 1. Imports
// 2. Types/interfaces
// 3. Component (export default for pages, named export for components)
// 4. Sub-components (if small/private)
```

## Styling
- Use Tailwind CSS v4 utilities directly
- Use `cn()` from `@/lib/utils` for conditional classes
- Prefer responsive utilities over media queries
- Use CSS variables via `@theme` for custom design tokens

## Existing Component Patterns
- `AuthGate.tsx` -- Auth wrapper (client component)
- `BayDiagram.tsx` -- SVG visualization
- `ElevationEditor.tsx` -- Form-heavy editor
- `CostSummary.tsx` -- Read-only data display
- `ReportOptionsDialog.tsx` -- Modal dialog pattern
Follow these existing patterns when creating new components.
