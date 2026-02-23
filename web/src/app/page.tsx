'use client';

import { useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { db } from '@/lib/database';
import { Plus, Trash2, Building2, FolderOpen, Loader2, X, Brain, ChevronRight, Sparkles } from 'lucide-react';

export default function HomePage() {
  const router = useRouter();
  const [projects, setProjects] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [newProjectName, setNewProjectName] = useState('');
  const [creating, setCreating] = useState(false);
  const [deletingProject, setDeletingProject] = useState<string | null>(null);
  const [confirmDelete, setConfirmDelete] = useState<string | null>(null);

  useEffect(() => {
    loadProjects();
  }, []);

  async function loadProjects() {
    try {
      setLoading(true);
      const projectList = await db.getProjects();
      setProjects(projectList);
    } catch (error) {
      console.error('Failed to load projects:', error);
    } finally {
      setLoading(false);
    }
  }

  async function handleCreateProject() {
    const trimmed = newProjectName.trim();
    if (!trimmed) return;

    if (projects.includes(trimmed)) {
      alert('A project with this name already exists.');
      return;
    }

    try {
      setCreating(true);
      await db.createProject(trimmed);
      setNewProjectName('');
      setShowCreateModal(false);
      await loadProjects();
    } catch (error) {
      console.error('Failed to create project:', error);
      alert('Failed to create project. Please try again.');
    } finally {
      setCreating(false);
    }
  }

  async function handleDeleteProject(projectName: string) {
    try {
      setDeletingProject(projectName);
      await db.deleteProject(projectName);
      setConfirmDelete(null);
      await loadProjects();
    } catch (error) {
      console.error('Failed to delete project:', error);
      alert('Failed to delete project. Please try again.');
    } finally {
      setDeletingProject(null);
    }
  }

  function openProject(projectName: string) {
    router.push(`/project/${encodeURIComponent(projectName)}`);
  }

  return (
    <div className="min-h-screen">

      {/* Header */}
      <header className="glass-strong gradient-border-top border-b border-[#1e1e2a]/60 sticky top-0 z-30">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3.5">
            <div className="relative">
              <div className="absolute inset-0 rounded-xl bg-blue-500/20 blur-lg animate-breathe" />
              <div className="relative w-10 h-10 rounded-xl bg-gradient-to-br from-[#3b82f6] to-[#6366f1] flex items-center justify-center shadow-lg shadow-blue-500/25">
                <Building2 className="w-5 h-5 text-white" />
              </div>
            </div>
            <div>
              <h1 className="text-lg font-bold gradient-text-static tracking-tight">
                United Glass Ventures
              </h1>
              <p className="text-[10px] text-[#55566a] font-semibold tracking-[0.2em] uppercase">
                Estimator Pro
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2.5">
            <button
              onClick={() => router.push('/ml-analytics')}
              className="flex items-center gap-2 px-4 py-2.5 bg-gradient-to-r from-purple-600 to-indigo-600 hover:brightness-110 text-white text-sm font-semibold rounded-xl transition-colors duration-200"
            >
              <Brain className="w-4 h-4" />
              ML Analytics
            </button>
            <button
              onClick={() => setShowCreateModal(true)}
              className="flex items-center gap-2 px-4 py-2.5 bg-gradient-to-r from-[#3b82f6] to-[#6366f1] hover:brightness-110 text-white text-sm font-semibold rounded-xl transition-colors duration-200"
            >
              <Plus className="w-4 h-4" />
              New Project
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-6 py-8 relative z-10">
        {/* Loading State */}
        {loading && (
          <div className="flex flex-col items-center justify-center py-32">
            <Loader2 className="w-8 h-8 text-[#3b82f6] animate-spin mb-4" />
            <p className="text-[#8b8d9a] text-sm">Loading projects...</p>
          </div>
        )}

        {/* Empty State */}
        {!loading && projects.length === 0 && (
          <div className="flex flex-col items-center justify-center py-32 animate-fade-up opacity-0">
            <div className="relative mb-8">
              <div className="absolute inset-0 w-24 h-24 rounded-2xl bg-blue-500/8 blur-2xl animate-breathe" />
              <div className="relative w-24 h-24 rounded-2xl bg-[#111118] border border-[#1e1e2a] flex items-center justify-center shadow-2xl shadow-black/30">
                <FolderOpen className="w-11 h-11 text-[#3e3f4d]" />
              </div>
            </div>
            <h2 className="text-2xl font-bold gradient-text-static tracking-tight mb-3">
              No projects yet
            </h2>
            <p className="text-[#8b8d9a] text-sm mb-10 max-w-sm text-center leading-relaxed">
              Get started by creating your first estimation project. Each project can contain multiple elevations and cost breakdowns.
            </p>
            <button
              onClick={() => setShowCreateModal(true)}
              className="flex items-center gap-2.5 px-7 py-3.5 bg-gradient-to-r from-[#3b82f6] to-[#6366f1] hover:brightness-110 text-white text-sm font-semibold rounded-xl transition-colors duration-200"
            >
              <Sparkles className="w-4 h-4" />
              Create Your First Project
            </button>
          </div>
        )}

        {/* Project Grid */}
        {!loading && projects.length > 0 && (
          <>
            <div className="mb-6 flex items-center justify-between">
              <h2 className="text-xs font-semibold text-[#55566a] uppercase tracking-[0.15em]">
                Projects ({projects.length})
              </h2>
              <div className="h-px flex-1 ml-4 bg-gradient-to-r from-[#1e1e2a] to-transparent" />
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3.5">
              {projects.map((project, index) => (
                <div
                  key={project}
                  className={`group relative card-hover bg-[#111118]/80 border border-[#1e1e2a] rounded-xl p-5 cursor-pointer transition-colors duration-200 hover:bg-[#13131b] animate-fade-up opacity-0 stagger-${Math.min(index + 1, 8)}`}
                  onClick={() => openProject(project)}
                >
                  <div className="flex items-start justify-between">
                    <div className="flex items-center gap-3.5 min-w-0">
                      <div className="relative">
                        <div className="absolute inset-0 rounded-xl bg-blue-500/15 blur-md opacity-0 group-hover:opacity-100 transition-opacity duration-400" />
                        <div className="relative w-11 h-11 rounded-xl bg-gradient-to-br from-[#3b82f6]/15 to-[#6366f1]/10 border border-[#3b82f6]/15 flex items-center justify-center flex-shrink-0">
                          <Building2 className="w-5 h-5 text-[#3b82f6] transition-colors duration-300 group-hover:text-[#60a5fa]" />
                        </div>
                      </div>
                      <div className="min-w-0">
                        <h3 className="text-sm font-semibold text-[#eeeef2] truncate group-hover:text-white transition-colors duration-200">
                          {project}
                        </h3>
                        <p className="text-xs text-[#3e3f4d] mt-0.5 flex items-center gap-1 transition-colors duration-200 group-hover:text-[#60a5fa]/60">
                          <span>Open project</span>
                          <ChevronRight className="w-3 h-3 transition-transform duration-300 group-hover:translate-x-1" />
                        </p>
                      </div>
                    </div>

                    {/* Delete Button */}
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        setConfirmDelete(project);
                      }}
                      className="opacity-0 group-hover:opacity-100 p-1.5 rounded-lg hover:bg-[#f87171]/10 text-[#3e3f4d] hover:text-[#f87171] transition-colors duration-200"
                      title="Delete project"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </>
        )}
      </main>

      {/* Create Project Modal */}
      {showCreateModal && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-md animate-overlay"
          onClick={() => {
            if (!creating) setShowCreateModal(false);
          }}
        >
          <div
            className="bg-[#111118] border border-[#1e1e2a] rounded-2xl w-full max-w-md mx-4 p-7 shadow-2xl shadow-black/60 animate-scale-in"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-6">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-lg bg-[#3b82f6]/10 border border-[#3b82f6]/20 flex items-center justify-center">
                  <Plus className="w-4 h-4 text-[#3b82f6]" />
                </div>
                <h2 className="text-lg font-semibold text-[#eeeef2] tracking-tight">
                  New Project
                </h2>
              </div>
              <button
                onClick={() => {
                  if (!creating) {
                    setShowCreateModal(false);
                    setNewProjectName('');
                  }
                }}
                className="p-1.5 rounded-lg hover:bg-[#1e1e2a] text-[#8b8d9a] hover:text-[#eeeef2] transition-colors duration-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="mb-6">
              <label
                htmlFor="project-name"
                className="block text-sm font-medium text-[#8b8d9a] mb-2"
              >
                Project Name
              </label>
              <input
                id="project-name"
                type="text"
                value={newProjectName}
                onChange={(e) => setNewProjectName(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleCreateProject();
                }}
                placeholder="e.g. Riverside Office Tower"
                className="w-full px-4 py-3 bg-[#0c0c12] border border-[#1e1e2a] rounded-xl text-sm text-[#eeeef2] placeholder-[#3e3f4d] focus:outline-none focus:border-[#3b82f6] focus:ring-2 focus:ring-[#3b82f6]/20 transition-colors duration-200"
                autoFocus
                disabled={creating}
              />
            </div>

            <div className="flex items-center gap-3 justify-end">
              <button
                onClick={() => {
                  if (!creating) {
                    setShowCreateModal(false);
                    setNewProjectName('');
                  }
                }}
                className="px-4 py-2.5 text-sm font-medium text-[#8b8d9a] hover:text-[#eeeef2] rounded-xl hover:bg-[#1e1e2a] transition-colors duration-200"
                disabled={creating}
              >
                Cancel
              </button>
              <button
                onClick={handleCreateProject}
                disabled={!newProjectName.trim() || creating}
                className="flex items-center gap-2 px-5 py-2.5 bg-gradient-to-r from-[#3b82f6] to-[#2563eb] hover:brightness-110 disabled:opacity-50 disabled:cursor-not-allowed text-white text-sm font-semibold rounded-xl transition-colors duration-200"
              >
                {creating ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    Creating...
                  </>
                ) : (
                  <>
                    <Plus className="w-4 h-4" />
                    Create Project
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete Confirmation Modal */}
      {confirmDelete && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-md animate-overlay"
          onClick={() => {
            if (!deletingProject) setConfirmDelete(null);
          }}
        >
          <div
            className="bg-[#111118] border border-[#1e1e2a] rounded-2xl w-full max-w-sm mx-4 p-7 shadow-2xl shadow-black/60 animate-scale-in"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center gap-3 mb-4">
              <div className="w-11 h-11 rounded-full bg-[#f87171]/10 border border-[#f87171]/15 flex items-center justify-center flex-shrink-0">
                <Trash2 className="w-5 h-5 text-[#f87171]" />
              </div>
              <div>
                <h3 className="text-base font-semibold text-[#eeeef2] tracking-tight">
                  Delete Project
                </h3>
                <p className="text-xs text-[#8b8d9a]">
                  This action cannot be undone
                </p>
              </div>
            </div>

            <p className="text-sm text-[#8b8d9a] mb-6 leading-relaxed">
              Are you sure you want to delete{' '}
              <span className="font-semibold text-[#eeeef2]">
                {confirmDelete}
              </span>
              ? All elevations and data will be permanently removed.
            </p>

            <div className="flex items-center gap-3 justify-end">
              <button
                onClick={() => {
                  if (!deletingProject) setConfirmDelete(null);
                }}
                className="px-4 py-2.5 text-sm font-medium text-[#8b8d9a] hover:text-[#eeeef2] rounded-xl hover:bg-[#1e1e2a] transition-colors duration-200"
                disabled={!!deletingProject}
              >
                Cancel
              </button>
              <button
                onClick={() => handleDeleteProject(confirmDelete)}
                disabled={!!deletingProject}
                className="flex items-center gap-2 px-4 py-2.5 bg-[#f87171] hover:bg-[#ef4444] disabled:opacity-50 disabled:cursor-not-allowed text-white text-sm font-semibold rounded-xl transition-colors duration-200"
              >
                {deletingProject === confirmDelete ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    Deleting...
                  </>
                ) : (
                  <>
                    <Trash2 className="w-4 h-4" />
                    Delete
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
