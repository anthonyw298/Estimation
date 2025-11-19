# Estimation Frontend

An Electron desktop application built with React, TypeScript, and Vite.

## Getting Started

### Prerequisites
- [Bun](https://bun.sh/) runtime installed

### Installation
```bash
bun install
```

### Development
Start the development server with hot reload:
```bash
bun dev
```

### Building
Build the application for production:
```bash
bun run prebuild
```

### Preview
Preview the production build:
```bash
bun start
```

## Build System & Tooling

This project uses **electron-vite** as the primary build tool, which orchestrates three separate Vite build processes:

1. **Main Process** (`electron/main/`) - Node.js environment running the Electron app lifecycle. Built as CommonJS with electron externalized.
2. **Preload Scripts** (`electron/preload/`) - Sandboxed scripts that expose safe APIs to the renderer via `contextBridge`.
3. **Renderer Process** (React app in `src/`) - Browser environment running the React UI with full HMR support.

**Key Technologies:**
- **Rolldown-Vite** - Fast Rust-based bundler (Vite 7 fork) for blazing build speeds
- **React Compiler** - Automatic optimization of React components (may impact dev performance)
- **Tailwind CSS v4** - Utility-first CSS via `@tailwindcss/vite` plugin
- **shadcn/ui** - Headless component library with Radix UI primitives
- **Bun** - JavaScript runtime and package manager for faster installs and execution

The `electron.vite.config.ts` defines entry points for all three processes. The main and preload scripts are compiled to `out/` directory, while the renderer dev server runs on `localhost:5173`. In production, the renderer HTML is bundled and served from the filesystem.

## Expanding the ESLint configuration

If you are developing a production application, we recommend updating the configuration to enable type-aware lint rules:

```js
export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...

      // Remove tseslint.configs.recommended and replace with this
      tseslint.configs.recommendedTypeChecked,
      // Alternatively, use this for stricter rules
      tseslint.configs.strictTypeChecked,
      // Optionally, add this for stylistic rules
      tseslint.configs.stylisticTypeChecked,

      // Other configs...
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```

You can also install [eslint-plugin-react-x](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-x) and [eslint-plugin-react-dom](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-dom) for React-specific lint rules:

```js
// eslint.config.js
import reactX from 'eslint-plugin-react-x'
import reactDom from 'eslint-plugin-react-dom'

export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...
      // Enable lint rules for React
      reactX.configs['recommended-typescript'],
      // Enable lint rules for React DOM
      reactDom.configs.recommended,
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```
