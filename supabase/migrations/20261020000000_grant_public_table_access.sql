-- Preserve the current app's public Data API behavior when new-table defaults
-- stop granting access automatically. The browser client uses the anon key for
-- reads and writes, and the existing RLS policies permit public access.
GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE
    public.projects,
    public.elevations,
    public.settings,
    public.doors,
    public.materials
TO anon, authenticated, service_role;

-- These tables use SERIAL primary keys; inserts also need sequence privileges.
GRANT USAGE, SELECT ON SEQUENCE
    public.projects_id_seq,
    public.elevations_id_seq,
    public.settings_id_seq,
    public.doors_id_seq,
    public.materials_id_seq
TO anon, authenticated, service_role;
