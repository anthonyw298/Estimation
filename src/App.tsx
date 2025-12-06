import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import ProjectsView from './views/ProjectsView';
import WorkspaceView from './views/WorkspaceView';
import './App.css';

function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<ProjectsView />} />
        <Route path="/workspace" element={<WorkspaceView />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </BrowserRouter>
  );
}

export default App;

