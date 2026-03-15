# Code Style

## TypeScript / React
- Use `const` by default, `let` only when reassignment is needed
- Prefer arrow functions for components and callbacks
- Destructure props in function parameters
- Use explicit return types on exported functions
- Prefer `interface` over `type` for object shapes
- Use `satisfies` for type-safe object literals

## Naming
- Components: PascalCase (`ElevationEditor.tsx`)
- Utilities/hooks: camelCase (`useProject.ts`, `formulas.ts`)
- Types/interfaces: PascalCase with no `I` prefix (`Project`, not `IProject`)
- Constants: UPPER_SNAKE_CASE for true constants, camelCase for config objects
- CSS classes: use Tailwind utilities, avoid custom class names

## Imports
- Use `@/` path alias for all `web/src/` imports
- Group: React/Next.js first, third-party second, local third
- No barrel exports -- import directly from the source file

## Python
- Follow PEP 8
- Use type hints for function signatures
- Docstrings for public functions
- f-strings over .format() or %
