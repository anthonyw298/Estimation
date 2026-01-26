-- Supabase Database Schema for EstimatorApp
-- Run this in Supabase SQL Editor: https://supabase.com → Your Project → SQL Editor

-- Projects table
CREATE TABLE IF NOT EXISTS projects (
    id SERIAL PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Elevations table
CREATE TABLE IF NOT EXISTS elevations (
    id SERIAL PRIMARY KEY,
    project_name TEXT NOT NULL,
    name TEXT NOT NULL,
    data JSONB DEFAULT '{}',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    UNIQUE(project_name, name)
);

-- Settings table
CREATE TABLE IF NOT EXISTS settings (
    id SERIAL PRIMARY KEY,
    project_name TEXT UNIQUE NOT NULL,
    data JSONB DEFAULT '{}',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Doors table
CREATE TABLE IF NOT EXISTS doors (
    id SERIAL PRIMARY KEY,
    project_name TEXT NOT NULL,
    elevation_name TEXT NOT NULL,
    data JSONB DEFAULT '[]',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    UNIQUE(project_name, elevation_name)
);

-- Materials table
CREATE TABLE IF NOT EXISTS materials (
    id SERIAL PRIMARY KEY,
    project_name TEXT UNIQUE NOT NULL,
    data JSONB DEFAULT '{}',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Enable Row Level Security (optional but recommended)
ALTER TABLE projects ENABLE ROW LEVEL SECURITY;
ALTER TABLE elevations ENABLE ROW LEVEL SECURITY;
ALTER TABLE settings ENABLE ROW LEVEL SECURITY;
ALTER TABLE doors ENABLE ROW LEVEL SECURITY;
ALTER TABLE materials ENABLE ROW LEVEL SECURITY;

-- Allow public access (for simplicity - you can add auth later)
CREATE POLICY "Allow all access to projects" ON projects FOR ALL USING (true);
CREATE POLICY "Allow all access to elevations" ON elevations FOR ALL USING (true);
CREATE POLICY "Allow all access to settings" ON settings FOR ALL USING (true);
CREATE POLICY "Allow all access to doors" ON doors FOR ALL USING (true);
CREATE POLICY "Allow all access to materials" ON materials FOR ALL USING (true);

-- Create indexes for faster queries
CREATE INDEX IF NOT EXISTS idx_elevations_project ON elevations(project_name);
CREATE INDEX IF NOT EXISTS idx_doors_project ON doors(project_name);
CREATE INDEX IF NOT EXISTS idx_doors_elevation ON doors(project_name, elevation_name);

