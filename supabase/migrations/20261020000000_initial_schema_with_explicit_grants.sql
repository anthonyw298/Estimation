-- Initial database schema for EstimatorApp.
-- Keep this aligned with ../../supabase_schema.sql.

CREATE TABLE IF NOT EXISTS public.projects (
    id SERIAL PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.projects
TO anon, authenticated, service_role;
GRANT USAGE, SELECT ON SEQUENCE public.projects_id_seq
TO anon, authenticated, service_role;

CREATE TABLE IF NOT EXISTS public.elevations (
    id SERIAL PRIMARY KEY,
    project_name TEXT NOT NULL,
    name TEXT NOT NULL,
    data JSONB DEFAULT '{}',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    UNIQUE(project_name, name)
);

GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.elevations
TO anon, authenticated, service_role;
GRANT USAGE, SELECT ON SEQUENCE public.elevations_id_seq
TO anon, authenticated, service_role;

CREATE TABLE IF NOT EXISTS public.settings (
    id SERIAL PRIMARY KEY,
    project_name TEXT UNIQUE NOT NULL,
    data JSONB DEFAULT '{}',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.settings
TO anon, authenticated, service_role;
GRANT USAGE, SELECT ON SEQUENCE public.settings_id_seq
TO anon, authenticated, service_role;

CREATE TABLE IF NOT EXISTS public.doors (
    id SERIAL PRIMARY KEY,
    project_name TEXT NOT NULL,
    elevation_name TEXT NOT NULL,
    data JSONB DEFAULT '[]',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    UNIQUE(project_name, elevation_name)
);

GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.doors
TO anon, authenticated, service_role;
GRANT USAGE, SELECT ON SEQUENCE public.doors_id_seq
TO anon, authenticated, service_role;

CREATE TABLE IF NOT EXISTS public.materials (
    id SERIAL PRIMARY KEY,
    project_name TEXT UNIQUE NOT NULL,
    data JSONB DEFAULT '{}',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.materials
TO anon, authenticated, service_role;
GRANT USAGE, SELECT ON SEQUENCE public.materials_id_seq
TO anon, authenticated, service_role;

ALTER TABLE public.projects ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.elevations ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.settings ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.doors ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.materials ENABLE ROW LEVEL SECURITY;

-- Preserve the current app policy for the existing five tables.
DROP POLICY IF EXISTS "Allow all access to projects" ON public.projects;
CREATE POLICY "Allow all access to projects" ON public.projects FOR ALL USING (true);
DROP POLICY IF EXISTS "Allow all access to elevations" ON public.elevations;
CREATE POLICY "Allow all access to elevations" ON public.elevations FOR ALL USING (true);
DROP POLICY IF EXISTS "Allow all access to settings" ON public.settings;
CREATE POLICY "Allow all access to settings" ON public.settings FOR ALL USING (true);
DROP POLICY IF EXISTS "Allow all access to doors" ON public.doors;
CREATE POLICY "Allow all access to doors" ON public.doors FOR ALL USING (true);
DROP POLICY IF EXISTS "Allow all access to materials" ON public.materials;
CREATE POLICY "Allow all access to materials" ON public.materials FOR ALL USING (true);

CREATE INDEX IF NOT EXISTS idx_elevations_project ON public.elevations(project_name);
CREATE INDEX IF NOT EXISTS idx_doors_project ON public.doors(project_name);
CREATE INDEX IF NOT EXISTS idx_doors_elevation ON public.doors(project_name, elevation_name);
