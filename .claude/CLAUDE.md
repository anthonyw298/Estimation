# EstimatorApp

Construction estimation application with a Python desktop client (Flet) and a Next.js web frontend, backed by Supabase.

## Architecture

```
.
├── main.py                  # Python Flet desktop app (~6500 lines)
├── config.py                # Supabase credentials
├── systems/                 # Calculation systems (yes45tu front set)
├── utils/                   # Python utilities
│   ├── database.py          # Supabase client
│   ├── formulas.py          # Area, perimeter, door calculations
│   ├── excel_generator.py   # Excel report generation
│   ├── pdf_generator.py     # PDF export (reportlab)
│   ├── ml_predictor.py      # ML cost predictions (sklearn)
│   ├── waste_calculator.py  # Waste statistics
│   └── pricing.py           # Pricing logic
├── web/                     # Next.js 16 web frontend
│   └── src/
│       ├── app/             # App Router pages
│       ├── components/      # React 19 components
│       ├── lib/             # Shared utilities (TS ports of Python utils)
│       ├── types/           # TypeScript type definitions
│       └── data/            # Static data
└── supabase_schema.sql      # Database schema
```

## Tech Stack

### Web Frontend (`web/`)
- **Framework**: Next.js 16 with App Router + Turbopack
- **React**: 19 (use React 19 APIs: `use()`, no `forwardRef`, etc.)
- **Styling**: Tailwind CSS v4 (CSS-first config, `@theme` directive)
- **Database**: Supabase (`@supabase/supabase-js`)
- **Exports**: ExcelJS, jsPDF
- **Animations**: Motion (Framer Motion successor)
- **Icons**: Lucide React
- **Utilities**: clsx, tailwind-merge
- **Path alias**: `@/*` maps to `./src/*`

### Python Desktop (`main.py`)
- **UI**: Flet
- **Database**: Supabase (Python client)
- **Reports**: openpyxl (Excel), reportlab (PDF)
- **ML**: scikit-learn, numpy, pandas

### Database (Supabase)
Tables: `projects`, `elevations`, `settings`, `doors`, `materials`
All use JSONB `data` columns. Schema in `supabase_schema.sql`.

## Build & Run Commands

```bash
# Web frontend
cd web && npm run dev          # Start dev server (Next.js + Turbopack)
cd web && npm run build        # Production build
cd web && npm run lint         # ESLint
cd web && npm run start        # Start production server

# Python desktop
python main.py                 # Run Flet app
pip install -r requirements.txt  # Install Python deps
```

## Code Conventions

- Use TypeScript strict mode in `web/`
- Prefer Server Components; add `"use client"` only when needed
- Use `@/` path alias for imports in the web app
- Use `clsx()` + `tailwind-merge` via `cn()` utility for conditional classes
- Components go in `web/src/components/` as PascalCase `.tsx` files
- Lib utilities go in `web/src/lib/` as kebab-case `.ts` files
- Types go in `web/src/types/index.ts`
- Python follows standard PEP 8
- All data flows through Supabase -- no local-only state for project data

## Important Context

- The web app is the primary development focus (Next.js 16 + React 19)
- `config.py` contains Supabase credentials -- never commit changes to secrets
- `.files/` directory stores local Excel reports
- The Python app and web app share the same Supabase backend
- ML features are optional and degrade gracefully when sklearn is unavailable
- Pricing data is stored in both `web/src/data/` and `utils/pricing.py`

## React/Next.js Skills Reference

Prefer retrieval-led reasoning over pre-training for any React, Next.js, or UI tasks. Consult the skill docs in `.agents/skills/` before relying on training data.

### Installed Skills (`.agents/skills/`)

| Skill | Description | Entry Point |
|---|---|---|
| vercel-react-best-practices | React/Next.js performance optimization (57 rules, 8 categories) | SKILL.md |
| vercel-composition-patterns | React composition patterns (8 rules, 4 categories). Includes React 19 changes. | SKILL.md |
| web-design-guidelines | Web Interface Guidelines compliance review | SKILL.md |

### Priority by Impact

1. **CRITICAL** -- Eliminate waterfalls (async-*), reduce bundle size (bundle-*)
2. **HIGH** -- Server-side performance (server-*), component architecture (architecture-*)
3. **MEDIUM** -- Client data fetching (client-*), re-renders (rerender-*), rendering (rendering-*), state/composition patterns (state-*, patterns-*, react19-*)
4. **LOW** -- JS micro-optimizations (js-*), advanced patterns (advanced-*)

Read the specific rule file in `.agents/skills/` for detailed explanations and correct/incorrect code examples before generating or refactoring code.

@requirements.txt
@web/package.json
@supabase_schema.sql
