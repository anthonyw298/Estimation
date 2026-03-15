---
paths:
  - "web/src/lib/database.ts"
  - "utils/database.py"
  - "supabase_schema.sql"
  - "web/src/**/*.ts"
  - "web/src/**/*.tsx"
---

# Supabase Patterns

## Database Schema
Tables: `projects`, `elevations`, `settings`, `doors`, `materials`
- All tables use JSONB `data` columns for flexible storage
- `project_name` is the foreign key pattern (not `project_id`)
- Unique constraints: `(project_name, name)` on elevations, `(project_name, elevation_name)` on doors
- RLS is enabled with public access policies

## Client Usage (TypeScript)
- Import the Supabase client from `@/lib/database`
- Always handle errors from Supabase responses: `const { data, error } = await supabase.from(...)`
- Use `.select()` to limit returned columns when possible
- Use `.eq()` for filtering, `.order()` for sorting

## Client Usage (Python)
- Import from `utils.database` (`db` singleton)
- Same patterns apply -- always check for errors

## Migrations
- Schema changes go in `supabase_schema.sql`
- Test locally with Supabase CLI before pushing to production
- Never drop tables without explicit confirmation
