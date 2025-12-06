import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { loadProjects, saveProjects, deleteProject } from '../utils/storage';
import '../App.css';

const COLOR_BG = '#000000';
const COLOR_SURFACE = '#1A1A1A';
const COLOR_ACCENT = '#0073E6';
const COLOR_TEXT = '#FFFFFF';
const COLOR_TEXT_DIM = '#B3B3B3';

export default function ProjectsView() {
  const [projects, setProjects] = useState<string[]>([]);
  const [newProjectName, setNewProjectName] = useState('');
  const [snackMessage, setSnackMessage] = useState('');
  const navigate = useNavigate();

  useEffect(() => {
    setProjects(loadProjects());
  }, []);

  const handleAddProject = () => {
    if (!newProjectName.trim()) return;
    if (projects.includes(newProjectName)) {
      showSnack('Project exists!');
      return;
    }
    const updated = [...projects, newProjectName];
    setProjects(updated);
    saveProjects(updated);
    setNewProjectName('');
    showSnack('Project created successfully');
  };

  const handleDeleteProject = (name: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (window.confirm(`Delete project "${name}"?`)) {
      deleteProject(name);
      const updated = projects.filter(p => p !== name);
      setProjects(updated);
      showSnack(`Project '${name}' deleted`);
    }
  };

  const handleProjectClick = (name: string) => {
    navigate('/workspace', { state: { projectName: name } });
  };

  const showSnack = (msg: string) => {
    setSnackMessage(msg);
    setTimeout(() => setSnackMessage(''), 3000);
  };

  return (
    <div style={{ backgroundColor: COLOR_BG, minHeight: '100vh', padding: '40px' }}>
      <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
        {/* Header */}
        <div style={{ display: 'flex', alignItems: 'center', marginBottom: '40px', gap: '25px' }}>
          <div style={{ 
            width: '160px', 
            height: '160px', 
            display: 'flex', 
            alignItems: 'center', 
            justifyContent: 'center',
            backgroundColor: COLOR_SURFACE,
            borderRadius: '10px',
            overflow: 'hidden'
          }}>
            <img 
              src="/assets/R.png" 
              alt="United Glass Logo" 
              style={{ 
                width: '100%', 
                height: '100%', 
                objectFit: 'contain',
                display: 'block'
              }}
              onError={(e) => {
                // Fallback to placeholder if image fails to load
                const img = e.currentTarget;
                img.style.display = 'none';
                const parent = img.parentElement;
                if (parent && !parent.querySelector('.logo-fallback')) {
                  const fallback = document.createElement('span');
                  fallback.className = 'logo-fallback';
                  fallback.style.cssText = 'font-size: 140px; font-weight: bold; color: #0073E6;';
                  fallback.textContent = 'U';
                  parent.appendChild(fallback);
                }
              }}
            />
          </div>
          <div style={{ flex: 1 }}>
            <h1 style={{ fontSize: '40px', fontWeight: 'bold', color: COLOR_ACCENT, marginBottom: '8px' }}>
              ESTIMATION TOOL
            </h1>
            <p style={{ fontSize: '18px', color: COLOR_TEXT_DIM }}>
              Select or create a project to begin
            </p>
          </div>
        </div>

        {/* New Project Input */}
        <div style={{ display: 'flex', gap: '10px', marginBottom: '20px' }}>
          <input
            type="text"
            placeholder="New Project Name"
            value={newProjectName}
            onChange={(e) => setNewProjectName(e.target.value)}
            onKeyPress={(e) => e.key === 'Enter' && handleAddProject()}
            className="input-field"
            style={{ flex: 1 }}
          />
          <button onClick={handleAddProject} className="btn btn-primary" style={{ fontSize: '40px', padding: '0 20px' }}>
            +
          </button>
        </div>

        {/* Projects Grid */}
        <div style={{ 
          display: 'grid', 
          gridTemplateColumns: 'repeat(auto-fill, minmax(160px, 1fr))', 
          gap: '20px',
          marginTop: '20px'
        }}>
          {projects.map((project) => (
            <div
              key={project}
              onClick={() => handleProjectClick(project)}
              style={{
                width: '160px',
                height: '160px',
                backgroundColor: COLOR_SURFACE,
                borderRadius: '10px',
                padding: '10px',
                cursor: 'pointer',
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                justifyContent: 'center',
                position: 'relative',
                transition: 'transform 0.2s',
              }}
              onMouseEnter={(e) => e.currentTarget.style.transform = 'scale(1.05)'}
              onMouseLeave={(e) => e.currentTarget.style.transform = 'scale(1)'}
            >
              <button
                onClick={(e) => handleDeleteProject(project, e)}
                style={{
                  position: 'absolute',
                  top: '5px',
                  right: '5px',
                  background: 'none',
                  border: 'none',
                  color: 'red',
                  cursor: 'pointer',
                  fontSize: '16px',
                  padding: '5px',
                }}
              >
                ×
              </button>
              <div style={{ fontSize: '40px', color: COLOR_ACCENT, marginBottom: '10px' }}>📁</div>
              <div style={{ 
                fontSize: '16px', 
                fontWeight: 'bold', 
                color: COLOR_TEXT,
                textAlign: 'center',
                wordBreak: 'break-word',
                overflow: 'hidden',
                textOverflow: 'ellipsis',
                maxWidth: '100%'
              }}>
                {project}
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Snackbar */}
      {snackMessage && (
        <div className="snackbar">{snackMessage}</div>
      )}
    </div>
  );
}

