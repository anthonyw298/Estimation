# Setup Instructions for shadcn/ui with Tailwind CSS v4

**Context:** Monorepo with `frontend/` using Vite 7, React 19, Bun, and Tailwind CSS v4 (via `@tailwindcss/vite`). Setting up shadcn/ui.

**Tasks:**
1. Fix bunx CLI issue preventing `bunx --bun shadcn@latest init` from running (PostCSS dependency resolution in temp directory).
2. Initialize shadcn/ui with Tailwind v4 support. If the CLI fails, manually create `components.json` with Tailwind v4-compatible settings.
3. Create `src/lib/utils.js` with `cn()` using `clsx` and `tailwind-merge`.
4. Ensure Vite config has `@` alias pointing to `./src` for shadcn imports.
5. Update `src/index.css` to use Tailwind v4 syntax (`@import "tailwindcss"`).
6. Configure for Electron: keep `base: './'` in Vite config for relative paths.
7. Test by adding a shadcn component (e.g., `bunx --bun shadcn@latest add button`).

**Dependencies already installed:** `class-variance-authority`, `clsx`, `tailwind-merge`, `lucide-react`, `@radix-ui/react-slot`.

