import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { motion, AnimatePresence } from 'framer-motion';
import { FiFolder, FiPlus, FiX, FiArrowRight } from 'react-icons/fi';
import { loadProjects, saveProjects, deleteProject } from '../utils/storage';
import '../App.css';

const COLOR_BG = '#000000';
const COLOR_SURFACE = '#1A1A1A';
const COLOR_ACCENT = '#0073E6';
const COLOR_TEXT = '#FFFFFF';
const COLOR_TEXT_DIM = '#B3B3B3';

const containerVariants = {
  hidden: { opacity: 0 },
  visible: {
    opacity: 1,
    transition: {
      staggerChildren: 0.1,
      delayChildren: 0.2
    }
  }
};

const itemVariants = {
  hidden: { opacity: 0, y: 20, scale: 0.9 },
  visible: {
    opacity: 1,
    y: 0,
    scale: 1,
    transition: {
      type: 'spring',
      stiffness: 100,
      damping: 15
    }
  }
};

const cardHoverVariants = {
  rest: { scale: 1, y: 0 },
  hover: {
    scale: 1.05,
    y: -8,
    transition: {
      type: 'spring',
      stiffness: 300,
      damping: 20
    }
  }
};

export default function ProjectsView() {
  const [projects, setProjects] = useState<string[]>([]);
  const [newProjectName, setNewProjectName] = useState('');
  const [snackMessage, setSnackMessage] = useState('');
  const [isLoading, setIsLoading] = useState(true);
  const navigate = useNavigate();

  useEffect(() => {
    const loaded = loadProjects();
    setProjects(loaded);
    setIsLoading(false);
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

  if (isLoading) {
    return (
      <div style={{ 
        display: 'flex', 
        justifyContent: 'center', 
        alignItems: 'center', 
        minHeight: '100vh',
        backgroundColor: COLOR_BG
      }}>
        <motion.div
          animate={{ rotate: 360 }}
          transition={{ duration: 1, repeat: Infinity, ease: 'linear' }}
          style={{
            width: '50px',
            height: '50px',
            border: '4px solid rgba(0, 115, 230, 0.2)',
            borderTop: '4px solid #0073E6',
            borderRadius: '50%'
          }}
        />
      </div>
    );
  }

  return (
    <motion.div
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      transition={{ duration: 0.5 }}
      style={{ 
        backgroundColor: COLOR_BG, 
        minHeight: '100vh', 
        padding: '40px',
        background: 'linear-gradient(135deg, #000000 0%, #0a0a0a 100%)'
      }}
    >
      <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
        {/* Header */}
        <motion.div
          initial={{ opacity: 0, y: -30 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, ease: 'easeOut' }}
          style={{ 
            display: 'flex', 
            alignItems: 'center', 
            marginBottom: '50px', 
            gap: '30px' 
          }}
        >
          <motion.div
            whileHover={{ scale: 1.1, rotate: 5 }}
            whileTap={{ scale: 0.95 }}
            style={{ 
              width: '180px', 
              height: '180px', 
              display: 'flex', 
              alignItems: 'center', 
              justifyContent: 'center',
              background: 'linear-gradient(135deg, #1A1A1A 0%, #2A2A2A 100%)',
              borderRadius: '20px',
              overflow: 'hidden',
              boxShadow: '0 8px 24px rgba(0, 115, 230, 0.3)',
              border: '2px solid rgba(0, 115, 230, 0.3)'
            }}
          >
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
          </motion.div>
          <motion.div
            initial={{ opacity: 0, x: -20 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ delay: 0.2, duration: 0.6 }}
            style={{ flex: 1 }}
          >
            <motion.h1
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              transition={{ delay: 0.3 }}
              style={{ 
                fontSize: '48px', 
                fontWeight: 'bold', 
                background: 'linear-gradient(135deg, #0073E6 0%, #00A3FF 100%)',
                WebkitBackgroundClip: 'text',
                WebkitTextFillColor: 'transparent',
                marginBottom: '12px',
                letterSpacing: '2px'
              }}
            >
              ESTIMATION TOOL
            </motion.h1>
            <motion.p
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              transition={{ delay: 0.4 }}
              style={{ fontSize: '20px', color: COLOR_TEXT_DIM }}
            >
              Select or create a project to begin
            </motion.p>
          </motion.div>
        </motion.div>

        {/* New Project Input */}
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.3, duration: 0.5 }}
          style={{ display: 'flex', gap: '12px', marginBottom: '30px' }}
        >
          <motion.input
            whileFocus={{ scale: 1.02 }}
            type="text"
            placeholder="New Project Name"
            value={newProjectName}
            onChange={(e) => setNewProjectName(e.target.value)}
            onKeyPress={(e) => e.key === 'Enter' && handleAddProject()}
            className="input-field"
            style={{ flex: 1, fontSize: '16px' }}
          />
          <motion.button
            whileHover={{ scale: 1.1, rotate: 90 }}
            whileTap={{ scale: 0.9 }}
            onClick={handleAddProject}
            className="btn btn-primary"
            style={{ 
              fontSize: '32px', 
              padding: '0 24px',
              minWidth: '60px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}
          >
            <FiPlus />
          </motion.button>
        </motion.div>

        {/* Projects Grid */}
        <AnimatePresence mode="popLayout">
          {projects.length > 0 ? (
            <motion.div
              variants={containerVariants}
              initial="hidden"
              animate="visible"
              style={{ 
                display: 'grid', 
                gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))', 
                gap: '24px',
                marginTop: '30px'
              }}
            >
              {projects.map((project, index) => (
                <motion.div
                  key={project}
                  variants={itemVariants}
                  layout
                  initial="rest"
                  whileHover="hover"
                  whileTap={{ scale: 0.95 }}
                  onClick={() => handleProjectClick(project)}
                  style={{
                    width: '100%',
                    aspectRatio: '1',
                    background: 'linear-gradient(135deg, #1A1A1A 0%, #2A2A2A 100%)',
                    borderRadius: '16px',
                    padding: '24px',
                    cursor: 'pointer',
                    display: 'flex',
                    flexDirection: 'column',
                    alignItems: 'center',
                    justifyContent: 'center',
                    position: 'relative',
                    boxShadow: '0 4px 12px rgba(0, 115, 230, 0.2)',
                    border: '2px solid rgba(0, 115, 230, 0.2)',
                    overflow: 'hidden'
                  }}
                >
                  <motion.button
                    onClick={(e) => handleDeleteProject(project, e)}
                    whileHover={{ scale: 1.2, rotate: 90 }}
                    whileTap={{ scale: 0.9 }}
                    style={{
                      position: 'absolute',
                      top: '12px',
                      right: '12px',
                      background: 'rgba(220, 53, 69, 0.2)',
                      border: 'none',
                      color: '#dc3545',
                      cursor: 'pointer',
                      fontSize: '20px',
                      padding: '8px',
                      borderRadius: '8px',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      zIndex: 10
                    }}
                  >
                    <FiX />
                  </motion.button>
                  
                  <motion.div
                    initial={{ scale: 0 }}
                    animate={{ scale: 1 }}
                    transition={{ delay: index * 0.1, type: 'spring', stiffness: 200 }}
                    style={{ 
                      fontSize: '56px', 
                      color: COLOR_ACCENT, 
                      marginBottom: '16px',
                      filter: 'drop-shadow(0 0 10px rgba(0, 115, 230, 0.5))'
                    }}
                  >
                    <FiFolder />
                  </motion.div>
                  
                  <motion.div
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: index * 0.1 + 0.2 }}
                    style={{ 
                      fontSize: '16px', 
                      fontWeight: '600', 
                      color: COLOR_TEXT,
                      textAlign: 'center',
                      wordBreak: 'break-word',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      maxWidth: '100%',
                      lineHeight: '1.4'
                    }}
                  >
                    {project}
                  </motion.div>

                  <motion.div
                    initial={{ opacity: 0, x: -10 }}
                    whileHover={{ opacity: 1, x: 0 }}
                    style={{
                      position: 'absolute',
                      bottom: '16px',
                      right: '16px',
                      color: COLOR_ACCENT,
                      fontSize: '20px'
                    }}
                  >
                    <FiArrowRight />
                  </motion.div>
                </motion.div>
              ))}
            </motion.div>
          ) : (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              style={{
                textAlign: 'center',
                padding: '60px 20px',
                color: COLOR_TEXT_DIM
              }}
            >
              <motion.div
                animate={{ scale: [1, 1.1, 1] }}
                transition={{ repeat: Infinity, duration: 2 }}
                style={{ fontSize: '64px', marginBottom: '20px' }}
              >
                <FiFolder />
              </motion.div>
              <p style={{ fontSize: '18px' }}>No projects yet. Create your first project above!</p>
            </motion.div>
          )}
        </AnimatePresence>
      </div>

      {/* Snackbar */}
      <AnimatePresence>
        {snackMessage && (
          <motion.div
            initial={{ opacity: 0, y: 100, scale: 0.8 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 100, scale: 0.8 }}
            transition={{ type: 'spring', stiffness: 300, damping: 25 }}
            className="snackbar"
          >
            {snackMessage}
          </motion.div>
        )}
      </AnimatePresence>
    </motion.div>
  );
}
