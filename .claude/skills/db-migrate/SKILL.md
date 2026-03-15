---
name: db-migrate
description: Generate or review Supabase schema migrations
disable-model-invocation: true
allowed-tools: Read, Edit, Write, Grep
---

# Database Migration

Generate or review a Supabase schema migration.

## Context

Current schema is in `supabase_schema.sql`. Tables:
- `projects` (id, name, created_at, updated_at)
- `elevations` (id, project_name, name, data JSONB)
- `settings` (id, project_name, data JSONB)
- `doors` (id, project_name, elevation_name, data JSONB)
- `materials` (id, project_name, data JSONB)

## Steps

1. Read the current schema: `supabase_schema.sql`

2. Based on `$ARGUMENTS`, generate the migration SQL:
   - Use `IF NOT EXISTS` for safety
   - Add appropriate indexes
   - Update RLS policies if needed
   - Include both up and down migrations as comments

3. Append the migration to `supabase_schema.sql` with a dated comment header.

4. If the change affects the TypeScript types, update `web/src/types/index.ts`.

5. If the change affects database queries, update:
   - `web/src/lib/database.ts` (TypeScript client)
   - `utils/database.py` (Python client)

6. Warn before any destructive operations (DROP, DELETE, TRUNCATE).
