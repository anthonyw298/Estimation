'use client';

import { useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { db } from '@/lib/database';
import { Plus, Trash2, Building2, FolderOpen, Loader2, X, Brain } from 'lucide-react';

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
    <div className="min-h-screen bg-[#06060a]">
      {/* Big greeting */}
      <div className="text-center py-12">
        <h1 className="text-7xl font-extrabold text-white">Hi</h1>
      </div>

      {/* Header */}
      <header className="glass gradient-border-top border-b border-[#1e1e2a] bg-[#06060a]/80 backdrop-blur-sm sticky top-0 z-30">
        <div className="max-w-7xl mx-auto px-6 py-5 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-[#3b82f6] flex items-center justify-center">
              <Building2 className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-[#eeeef2] tracking-tight">
                United Glass Ventures
              </h1>
              <p className="text-xs text-[#55566a] font-medium tracking-wide uppercase">
                Estimator Pro
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <button
              onClick={() => router.push('/ml-analytics')}
              className="flex items-center gap-2 px-4 py-2.5 bg-purple-600 hover:bg-purple-700 text-white text-sm font-medium rounded-lg transition-all duration-200 active:scale-[0.97] shadow-md shadow-purple-500/10"
            >
              <Brain className="w-4 h-4" />
              ML Analytics
            </button>
            <button
              onClick={() => setShowCreateModal(true)}
              className="flex items-center gap-2 px-4 py-2.5 bg-[#3b82f6] hover:bg-[#2563eb] text-white text-sm font-medium rounded-lg transition-all duration-200 active:scale-[0.97] shadow-md shadow-blue-500/10"
            >
              <Plus className="w-4 h-4" />
              New Project
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-6 py-8">
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
            <div className="w-16 h-16 rounded-2xl bg-[#111118] border border-[#1e1e2a] flex items-center justify-center mb-6">
              <FolderOpen className="w-8 h-8 text-[#55566a]" />
            </div>
            <h2 className="text-xl font-semibold text-[#eeeef2] tracking-tight mb-2">
              No projects yet
            </h2>
            <p className="text-[#8b8d9a] text-sm mb-6 max-w-sm text-center">
              Get started by creating your first estimation project. Each project can contain multiple elevations and cost breakdowns.
            </p>
            <button
              onClick={() => setShowCreateModal(true)}
              className="flex items-center gap-2 px-5 py-2.5 bg-[#3b82f6] hover:bg-[#2563eb] text-white text-sm font-medium rounded-lg transition-all duration-200 active:scale-[0.97] shadow-md shadow-blue-500/10"
            >
              <Plus className="w-4 h-4" />
              Create Your First Project
            </button>
          </div>
        )}

        {/* Project Grid */}
        {!loading && projects.length > 0 && (
          <>
            <div className="mb-6">
              <h2 className="text-sm font-medium text-[#8b8d9a] uppercase tracking-wider">
                Projects ({projects.length})
              </h2>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {projects.map((project, index) => (
                <div
                  key={project}
                  className={`group relative glow-blue-hover bg-[#111118] border border-[#1e1e2a] rounded-xl p-5 cursor-pointer transition-all duration-300 ease-out hover:scale-[1.015] hover:border-[#2a2a3a] hover:bg-[#16161f] hover:shadow-lg hover:shadow-black/30 animate-fade-up opacity-0 stagger-${index + 1}`}
                  onClick={() => openProject(project)}
                >
                  <div className="flex items-start justify-between">
                    <div className="flex items-center gap-3 min-w-0">
                      <div className="w-9 h-9 rounded-lg bg-[#3b82f6]/10 border border-[#3b82f6]/20 flex items-center justify-center flex-shrink-0">
                        <Building2 className="w-4 h-4 text-[#3b82f6]" />
                      </div>
                      <div className="min-w-0">
                        <h3 className="text-sm font-semibold text-[#eeeef2] truncate">
                          {project}
                        </h3>
                        <p className="text-xs text-[#55566a] mt-0.5">
                          Click to open
                        </p>
                      </div>
                    </div>

                    {/* Delete Button */}
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        setConfirmDelete(project);
                      }}
                      className="opacity-0 group-hover:opacity-100 p-1.5 rounded-md hover:bg-[#f87171]/10 text-[#55566a] hover:text-[#f87171] transition-all duration-200"
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
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm animate-overlay"
          onClick={() => {
            if (!creating) setShowCreateModal(false);
          }}
        >
          <div
            className="bg-[#111118] border border-[#1e1e2a] rounded-2xl w-full max-w-md mx-4 p-6 shadow-2xl shadow-black/50 animate-scale-in"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-6">
              <h2 className="text-lg font-semibold text-[#eeeef2] tracking-tight">
                New Project
              </h2>
              <button
                onClick={() => {
                  if (!creating) {
                    setShowCreateModal(false);
                    setNewProjectName('');
                  }
                }}
                className="p-1.5 rounded-md hover:bg-[#1e1e2a] text-[#8b8d9a] hover:text-[#eeeef2] transition-colors"
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
                className="w-full px-3.5 py-2.5 bg-[#0c0c12] border border-[#1e1e2a] rounded-lg text-sm text-[#eeeef2] placeholder-[#55566a] focus:outline-none focus:border-[#3b82f6] focus:ring-2 focus:ring-[#3b82f6]/20 transition-colors"
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
                className="px-4 py-2 text-sm font-medium text-[#8b8d9a] hover:text-[#eeeef2] rounded-lg hover:bg-[#1e1e2a] transition-all duration-200"
                disabled={creating}
              >
                Cancel
              </button>
              <button
                onClick={handleCreateProject}
                disabled={!newProjectName.trim() || creating}
                className="flex items-center gap-2 px-4 py-2 bg-[#3b82f6] hover:bg-[#2563eb] disabled:opacity-50 disabled:cursor-not-allowed text-white text-sm font-medium rounded-lg transition-all duration-200 active:scale-[0.97] shadow-md shadow-blue-500/10"
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
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm animate-overlay"
          onClick={() => {
            if (!deletingProject) setConfirmDelete(null);
          }}
        >
          <div
            className="bg-[#111118] border border-[#1e1e2a] rounded-2xl w-full max-w-sm mx-4 p-6 shadow-2xl shadow-black/50 animate-scale-in"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 rounded-full bg-[#f87171]/10 flex items-center justify-center flex-shrink-0">
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

            <p className="text-sm text-[#8b8d9a] mb-6">
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
                className="px-4 py-2 text-sm font-medium text-[#8b8d9a] hover:text-[#eeeef2] rounded-lg hover:bg-[#1e1e2a] transition-all duration-200"
                disabled={!!deletingProject}
              >
                Cancel
              </button>
              <button
                onClick={() => handleDeleteProject(confirmDelete)}
                disabled={!!deletingProject}
                className="flex items-center gap-2 px-4 py-2 bg-[#f87171] hover:bg-[#ef4444] disabled:opacity-50 disabled:cursor-not-allowed text-white text-sm font-medium rounded-lg transition-all duration-200 active:scale-[0.97] shadow-md shadow-red-500/10"
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
