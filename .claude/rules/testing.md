# Testing

## Web Frontend
- Run `npm run lint` before committing changes to `web/`
- Run `npm run build` to verify no TypeScript or build errors
- Test Server Components by checking they render without client-side APIs
- Test Client Components by verifying interactive behavior

## Python
- Test calculation functions in `utils/formulas.py` with known inputs
- Test database operations against Supabase
- ML predictor tests should handle missing sklearn gracefully

## Pre-commit Checklist
1. `cd web && npm run lint` -- no ESLint errors
2. `cd web && npm run build` -- build succeeds
3. No secrets or `.env` files in the commit
4. Types are correct (no `any` unless unavoidable)
