import { useState, useEffect } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import { motion, AnimatePresence } from 'framer-motion';
import { 
  FiArrowLeft, 
  FiFileText, 
  FiSave, 
  FiTrash2, 
  FiEdit, 
  FiX, 
  FiPlus,
  FiCheck,
  FiDownload,
  FiSettings
} from 'react-icons/fi';
import {
  loadElevations,
  saveElevations,
  loadDoors,
  saveDoors,
  loadExtraMaterials,
  ElevationData
} from '../utils/storage';
import { calculateRectangleArea, calculatePerimeter, calculateDoorInfo, DoorInfo } from '../utils/formulas';
import { calculateYes45tuQuantities } from '../systems/yes45tuFrontSet';
import { generateExcelReport } from '../utils/excelGenerator';
import '../App.css';

const COLOR_BG = '#000000';
const COLOR_SURFACE = '#1A1A1A';
const COLOR_ACCENT = '#0073E6';
const COLOR_TEXT = '#FFFFFF';
const COLOR_TEXT_DIM = '#B3B3B3';
const COLOR_INPUT_BG = '#2A2A2A';

const slideIn = {
  hidden: { opacity: 0, x: -20 },
  visible: { 
    opacity: 1, 
    x: 0,
    transition: { duration: 0.4, ease: 'easeOut' }
  },
  exit: { opacity: 0, x: 20, transition: { duration: 0.3 } }
};

const fadeIn = {
  hidden: { opacity: 0, y: 20 },
  visible: { 
    opacity: 1, 
    y: 0,
    transition: { duration: 0.4 }
  }
};

export default function WorkspaceView() {
  const location = useLocation();
  const navigate = useNavigate();
  const projectName = (location.state as any)?.projectName || '';

  const [elevations, setElevations] = useState<Record<string, ElevationData>>({});
  const [currentElevation, setCurrentElevation] = useState<string>('New Elevation');
  const [snackMessage, setSnackMessage] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  // Form state
  const [system, setSystem] = useState('YES 45TU FRONT SET(OG)');
  const [finish, setFinish] = useState('Clear');
  const [elevType, setElevType] = useState('');
  const [count, setCount] = useState('');
  const [width, setWidth] = useState('');
  const [height, setHeight] = useState('');
  const [baysWide, setBaysWide] = useState('');
  const [baysTall, setBaysTall] = useState('');
  const [customBayWidths, setCustomBayWidths] = useState<string[]>([]);
  const [customBayHeights, setCustomBayHeights] = useState<string[]>([]);

  // Door state
  const [doors, setDoors] = useState<DoorInfo[]>([]);
  const [doorSize, setDoorSize] = useState("3' X 7'");
  const [doorCount, setDoorCount] = useState('');
  const [doorStile, setDoorStile] = useState('Narrow');
  const [doorHardware, setDoorHardware] = useState<Record<string, boolean>>({
    'Continuous Hinges': false,
    'Concealed Closer': false,
    'Exit Devices': false,
    'Electric Strike': false,
    'Extended Ladder Pull (B2B)': false,
    'Extended Ladder Pull (Single)': false,
    'Latch Lock w/ Lever Handle': false,
    'Lever Handle': false
  });
  const [selectedDoorIndex, setSelectedDoorIndex] = useState<number | null>(null);

  const systemOptions = ['YES 45TU FRONT SET(OG)', 'Other'];
  const finishOptions = ['Clear', 'Black', 'Paint'];
  const doorOptions = ['None', "3' X 7'", "3' X 8'", "3' X 9'", "6' X 7'", "6' X 8'", "6' X 9'"];
  const stileOptions = ['Narrow', 'Medium', 'Wide'];
  const hardwareOptions = [
    'Continuous Hinges', 'Concealed Closer', 'Exit Devices', 'Electric Strike',
    'Extended Ladder Pull (B2B)', 'Extended Ladder Pull (Single)',
    'Latch Lock w/ Lever Handle', 'Lever Handle'
  ];

  useEffect(() => {
    if (!projectName) {
      navigate('/');
      return;
    }
    const elevs = loadElevations(projectName);
    setElevations(elevs);
    if (Object.keys(elevs).length > 0) {
      setCurrentElevation(Object.keys(elevs)[0]);
    }
  }, [projectName, navigate]);

  useEffect(() => {
    if (currentElevation && currentElevation !== 'New Elevation' && elevations[currentElevation]) {
      const elev = elevations[currentElevation];
      setSystem(elev.system || systemOptions[0]);
      setFinish(elev.finish || finishOptions[0]);
      setElevType(currentElevation);
      setCount(elev.total_count?.toString() || '');
      setWidth(elev.opening_width_inches?.toString() || '');
      setHeight(elev.opening_height_inches?.toString() || '');
      setBaysWide(elev.bays_wide?.toString() || '');
      setBaysTall(elev.bays_tall?.toString() || '');
      setCustomBayWidths(elev.custom_bay_widths?.map(w => w.toString()) || []);
      setCustomBayHeights(elev.custom_bay_heights?.map(h => h.toString()) || []);
      const loadedDoors = loadDoors(projectName, currentElevation);
      setDoors(loadedDoors);
    } else {
      clearWorkspace();
    }
  }, [currentElevation]);

  useEffect(() => {
    if (system === 'YES 45TU FRONT SET(OG)') {
      updateDynamicBayInputs();
    }
  }, [baysWide, baysTall, system]);

  const showSnack = (msg: string) => {
    setSnackMessage(msg);
    setTimeout(() => setSnackMessage(''), 3000);
  };

  const clearWorkspace = () => {
    setElevType('');
    setCount('');
    setWidth('');
    setHeight('');
    setBaysWide('');
    setBaysTall('');
    setCustomBayWidths([]);
    setCustomBayHeights([]);
    setDoors([]);
    setSelectedDoorIndex(null);
  };

  const updateDynamicBayInputs = () => {
    const bw = parseInt(baysWide) || 0;
    const bh = parseInt(baysTall) || 0;
    
    if (bw > 0 && customBayWidths.length !== bw) {
      setCustomBayWidths(Array(bw).fill(''));
    }
    if (bh > 0 && customBayHeights.length !== bh) {
      setCustomBayHeights(Array(bh).fill(''));
    }
  };

  const handleElevationLoad = (elevName: string) => {
    if (elevName === 'New Elevation') {
      clearWorkspace();
      setCurrentElevation('New Elevation');
      return;
    }
    setCurrentElevation(elevName);
  };

  const handleSaveElevation = async () => {
    try {
      setIsSaving(true);
      if (!elevType.trim()) throw new Error('Elevation Name Required');
      if (!count) throw new Error('Quantity is required');
      const total = parseInt(count);
      if (!width) throw new Error('Opening Width is required');
      const w = parseFloat(width);
      if (!height) throw new Error('Opening Height is required');
      const h = parseFloat(height);

      const sqft = calculateRectangleArea(w / 12, h / 12);
      const perim = calculatePerimeter(w / 12, h / 12);

      const data: ElevationData = {
        system: system,
        finish: finish,
        total_count: total,
        opening_width_inches: w,
        opening_height_inches: h,
        sqft_per_type: sqft,
        total_sqft: sqft * total,
        perimeter_ft: perim,
        total_perimeter_ft: perim * total,
        calculated_outputs: [],
        material_impact: []
      };

      if (system === 'YES 45TU FRONT SET(OG)') {
        if (!baysWide) throw new Error('Bays Wide is required for YES 45TU FRONT SET(OG)');
        if (!baysTall) throw new Error('Bays Tall is required for YES 45TU FRONT SET(OG)');
        const bw = parseInt(baysWide);
        const bh = parseInt(baysTall);

        const customW = customBayWidths.map(w => parseFloat(w) || 0).filter(w => w > 0);
        const customH = customBayHeights.map(h => parseFloat(h) || 0).filter(h => h > 0);

        data.bays_wide = bw;
        data.bays_tall = bh;
        data.custom_bay_widths = customW.length === bw ? customW : undefined;
        data.custom_bay_heights = customH.length === bh ? customH : undefined;

        const calculatedOutputs = calculateYes45tuQuantities(
          bw, bh, total, w, h, doors,
          data.custom_bay_widths
        );
        data.calculated_outputs = calculatedOutputs;
      }

      const doorItems = calculateDoorInfo(doors, finish);
      data.calculated_outputs.push(...doorItems);

      const updated = { ...elevations, [elevType]: data };
      setElevations(updated);
      saveElevations(projectName, updated);
      saveDoors(projectName, elevType, doors);

      showSnack('Elevation Saved Successfully');
      setCurrentElevation(elevType);
    } catch (e: any) {
      showSnack(`Error: ${e.message}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleDeleteElevation = () => {
    if (currentElevation && currentElevation !== 'New Elevation' && elevations[currentElevation]) {
      if (window.confirm(`Delete elevation "${currentElevation}"?`)) {
        const updated = { ...elevations };
        delete updated[currentElevation];
        setElevations(updated);
        saveElevations(projectName, updated);
        setCurrentElevation('New Elevation');
        clearWorkspace();
        showSnack('Elevation Deleted');
      }
    }
  };

  const handleAddDoor = () => {
    if (!doorCount) {
      showSnack('Invalid door count');
      return;
    }
    const newDoor: DoorInfo = {
      size: doorSize,
      count: parseInt(doorCount),
      stile: doorStile,
      hardware: { ...doorHardware }
    };
    setDoors([...doors, newDoor]);
    saveDoors(projectName, elevType || currentElevation, [...doors, newDoor]);
    setDoorCount('');
    setDoorHardware(Object.fromEntries(hardwareOptions.map(h => [h, false])));
    showSnack('Door Added');
  };

  const handleUpdateDoor = () => {
    if (selectedDoorIndex === null || !doorCount) {
      showSnack('Please select a door to update');
      return;
    }
    const updated = [...doors];
    updated[selectedDoorIndex] = {
      size: doorSize,
      count: parseInt(doorCount),
      stile: doorStile,
      hardware: { ...doorHardware }
    };
    setDoors(updated);
    saveDoors(projectName, elevType || currentElevation, updated);
    setSelectedDoorIndex(null);
    setDoorCount('');
    setDoorHardware(Object.fromEntries(hardwareOptions.map(h => [h, false])));
    showSnack('Door Updated');
  };

  const handleDeleteDoor = (index: number) => {
    const updated = doors.filter((_, i) => i !== index);
    setDoors(updated);
    saveDoors(projectName, elevType || currentElevation, updated);
    showSnack('Door Deleted');
  };

  const handleEditDoor = (index: number) => {
    const door = doors[index];
    setDoorSize(door.size);
    setDoorCount(door.count.toString());
    setDoorStile(door.stile);
    setDoorHardware({ ...door.hardware });
    setSelectedDoorIndex(index);
  };

  const autoFillWidths = () => {
    const totalW = parseFloat(width);
    if (!totalW) {
      showSnack('Please set valid Opening Width first');
      return;
    }
    const filled = customBayWidths.reduce((sum, w) => sum + (parseFloat(w) || 0), 0);
    const blankCount = customBayWidths.filter(w => !w || parseFloat(w) === 0).length;
    if (filled > totalW) {
      showSnack(`Error: Filled widths (${filled.toFixed(2)}) exceed total (${totalW.toFixed(2)})`);
      return;
    }
    if (blankCount > 0) {
      const remaining = totalW - filled;
      const share = remaining / blankCount;
      const updated = customBayWidths.map(w => w && parseFloat(w) > 0 ? w : share.toFixed(4));
      setCustomBayWidths(updated);
      showSnack('Auto-fill complete');
    }
  };

  const autoFillHeights = () => {
    const totalH = parseFloat(height);
    if (!totalH) {
      showSnack('Please set valid Opening Height first');
      return;
    }
    const filled = customBayHeights.reduce((sum, h) => sum + (parseFloat(h) || 0), 0);
    const blankCount = customBayHeights.filter(h => !h || parseFloat(h) === 0).length;
    if (filled > totalH) {
      showSnack(`Error: Filled heights (${filled.toFixed(2)}) exceed total (${totalH.toFixed(2)})`);
      return;
    }
    if (blankCount > 0) {
      const remaining = totalH - filled;
      const share = remaining / blankCount;
      const updated = customBayHeights.map(h => h && parseFloat(h) > 0 ? h : share.toFixed(4));
      setCustomBayHeights(updated);
      showSnack('Auto-fill complete');
    }
  };

  const handleGenerateReport = async () => {
    setIsGenerating(true);
    try {
      if (!projectName) {
        showSnack('No project selected');
        return;
      }
      
      const latestElevations = loadElevations(projectName);
      
      if (!latestElevations || Object.keys(latestElevations).length === 0) {
        showSnack('No elevations to generate report. Please create an elevation first.');
        return;
      }
      
      showSnack('Generating report...');
      await generateExcelReport(projectName, latestElevations);
      showSnack('Report generated successfully! Check your downloads folder.');
    } catch (e: any) {
      console.error('Report generation error:', e);
      showSnack(`Error generating report: ${e.message || e.toString()}`);
    } finally {
      setIsGenerating(false);
    }
  };

  const isYes45 = system === 'YES 45TU FRONT SET(OG)';
  const elevationNames = ['New Elevation', ...Object.keys(elevations).sort()];
  const isUpdateMode = currentElevation !== 'New Elevation' && elevations[currentElevation];

  return (
    <motion.div
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      transition={{ duration: 0.4 }}
      style={{ 
        backgroundColor: COLOR_BG, 
        minHeight: '100vh', 
        padding: '24px',
        background: 'linear-gradient(135deg, #000000 0%, #0a0a0a 100%)'
      }}
    >
      {/* Header */}
      <motion.div
        variants={fadeIn}
        initial="hidden"
        animate="visible"
        style={{ 
          display: 'flex', 
          alignItems: 'center', 
          marginBottom: '24px', 
          height: '60px',
          gap: '16px'
        }}
      >
        <motion.button
          whileHover={{ scale: 1.1, x: -4 }}
          whileTap={{ scale: 0.95 }}
          onClick={() => navigate('/')}
          className="icon-btn"
          style={{ fontSize: '24px' }}
        >
          <FiArrowLeft />
        </motion.button>
        
        <motion.h2
          initial={{ opacity: 0, x: -20 }}
          animate={{ opacity: 1, x: 0 }}
          transition={{ delay: 0.1 }}
          style={{ 
            color: COLOR_TEXT, 
            margin: 0,
            fontSize: '24px',
            fontWeight: 'bold',
            background: 'linear-gradient(135deg, #0073E6 0%, #00A3FF 100%)',
            WebkitBackgroundClip: 'text',
            WebkitTextFillColor: 'transparent'
          }}
        >
          {projectName.toUpperCase()}
        </motion.h2>
        
        <div style={{ flex: 1 }} />
        
        <motion.button
          whileHover={{ scale: 1.05 }}
          whileTap={{ scale: 0.95 }}
          className="btn btn-primary"
          type="button"
          onClick={handleGenerateReport}
          disabled={isGenerating}
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
            opacity: isGenerating ? 0.7 : 1
          }}
        >
          {isGenerating ? (
            <>
              <motion.div
                animate={{ rotate: 360 }}
                transition={{ duration: 1, repeat: Infinity, ease: 'linear' }}
                style={{
                  width: '16px',
                  height: '16px',
                  border: '2px solid rgba(255,255,255,0.3)',
                  borderTop: '2px solid white',
                  borderRadius: '50%'
                }}
              />
              Generating...
            </>
          ) : (
            <>
              <FiDownload />
              GENERATE REPORT
            </>
          )}
        </motion.button>
      </motion.div>

      <div style={{ display: 'flex', gap: '24px', flexWrap: 'wrap' }}>
        {/* Left Column: Elevation Form */}
        <motion.div
          variants={slideIn}
          initial="hidden"
          animate="visible"
          style={{ 
            flex: 2, 
            minWidth: '500px',
            background: 'linear-gradient(135deg, #1A1A1A 0%, #222222 100%)',
            borderRadius: '16px', 
            padding: '28px',
            boxShadow: '0 8px 24px rgba(0, 115, 230, 0.2)',
            border: '1px solid rgba(0, 115, 230, 0.2)'
          }}
        >
          <motion.select
            whileFocus={{ scale: 1.02 }}
            value={currentElevation}
            onChange={(e) => handleElevationLoad(e.target.value)}
            className="input-field"
            style={{ marginBottom: '24px', fontSize: '16px' }}
          >
            {elevationNames.map(name => (
              <option key={name} value={name}>{name}</option>
            ))}
          </motion.select>

          <motion.div
            variants={fadeIn}
            initial="hidden"
            animate="visible"
            className="form-section"
            style={{ marginBottom: '20px' }}
          >
            <label className="input-label">System</label>
            <select 
              value={system} 
              onChange={(e) => setSystem(e.target.value)} 
              className="input-field"
            >
              {systemOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
            </select>
          </motion.div>

          <motion.div
            variants={fadeIn}
            initial="hidden"
            animate="visible"
            className="form-section"
            style={{ marginBottom: '20px' }}
          >
            <label className="input-label">Finish</label>
            <select 
              value={finish} 
              onChange={(e) => setFinish(e.target.value)} 
              className="input-field"
            >
              {finishOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
            </select>
          </motion.div>

          <motion.div
            variants={fadeIn}
            initial="hidden"
            animate="visible"
            style={{ display: 'flex', gap: '12px', marginBottom: '20px' }}
          >
            <div style={{ flex: 1 }}>
              <label className="input-label">Elevation Type (Name)</label>
              <input 
                type="text" 
                value={elevType} 
                onChange={(e) => setElevType(e.target.value)} 
                className="input-field" 
              />
            </div>
            <div style={{ flex: 1 }}>
              <label className="input-label">Quantity</label>
              <input 
                type="number" 
                value={count} 
                onChange={(e) => setCount(e.target.value)} 
                className="input-field" 
              />
            </div>
          </motion.div>

          <motion.div
            variants={fadeIn}
            initial="hidden"
            animate="visible"
            className="form-section"
            style={{ marginTop: '24px', marginBottom: '20px' }}
          >
            <h3 style={{ color: COLOR_TEXT_DIM, fontSize: '14px', marginBottom: '12px', fontWeight: '600' }}>DIMENSIONS</h3>
            <div style={{ display: 'flex', gap: '12px' }}>
              <div style={{ flex: 1 }}>
                <label className="input-label">Opening Width (")</label>
                <input 
                  type="number" 
                  value={width} 
                  onChange={(e) => setWidth(e.target.value)} 
                  className="input-field" 
                />
              </div>
              <div style={{ flex: 1 }}>
                <label className="input-label">Opening Height (")</label>
                <input 
                  type="number" 
                  value={height} 
                  onChange={(e) => setHeight(e.target.value)} 
                  className="input-field" 
                />
              </div>
            </div>
          </motion.div>

          <AnimatePresence>
            {isYes45 && (
              <motion.div
                initial={{ opacity: 0, height: 0 }}
                animate={{ opacity: 1, height: 'auto' }}
                exit={{ opacity: 0, height: 0 }}
                transition={{ duration: 0.3 }}
                className="form-section"
                style={{ marginTop: '20px', marginBottom: '20px', overflow: 'hidden' }}
              >
                <h3 style={{ color: COLOR_TEXT_DIM, fontSize: '14px', marginBottom: '12px', fontWeight: '600' }}>BAY CONFIGURATION</h3>
                <div style={{ display: 'flex', gap: '12px', marginBottom: '16px' }}>
                  <div style={{ flex: 1 }}>
                    <label className="input-label">Bays Wide</label>
                    <input 
                      type="number" 
                      value={baysWide} 
                      onChange={(e) => setBaysWide(e.target.value)} 
                      className="input-field" 
                    />
                  </div>
                  <div style={{ flex: 1 }}>
                    <label className="input-label">Bays Tall</label>
                    <input 
                      type="number" 
                      value={baysTall} 
                      onChange={(e) => setBaysTall(e.target.value)} 
                      className="input-field" 
                    />
                  </div>
                </div>

                {parseInt(baysWide) > 0 && (
                  <motion.div
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    style={{ marginTop: '16px', marginBottom: '16px' }}
                  >
                    <label className="input-label">Custom Bay Widths (leave blank to auto-fill)</label>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '10px', marginBottom: '12px' }}>
                      {customBayWidths.map((w, i) => (
                        <motion.input
                          key={i}
                          whileFocus={{ scale: 1.05 }}
                          type="number"
                          value={w}
                          onChange={(e) => {
                            const updated = [...customBayWidths];
                            updated[i] = e.target.value;
                            setCustomBayWidths(updated);
                          }}
                          placeholder={`Bay ${i + 1}`}
                          className="input-field"
                        />
                      ))}
                    </div>
                    <motion.button
                      whileHover={{ scale: 1.02 }}
                      whileTap={{ scale: 0.98 }}
                      onClick={autoFillWidths}
                      className="btn btn-primary"
                      style={{ width: '100%' }}
                    >
                      Auto-Fill Remaining Widths
                    </motion.button>
                  </motion.div>
                )}

                {parseInt(baysTall) > 0 && (
                  <motion.div
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    style={{ marginTop: '16px', marginBottom: '16px' }}
                  >
                    <label className="input-label">Custom Bay Heights (leave blank to auto-fill)</label>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '10px', marginBottom: '12px' }}>
                      {customBayHeights.map((h, i) => (
                        <motion.input
                          key={i}
                          whileFocus={{ scale: 1.05 }}
                          type="number"
                          value={h}
                          onChange={(e) => {
                            const updated = [...customBayHeights];
                            updated[i] = e.target.value;
                            setCustomBayHeights(updated);
                          }}
                          placeholder={`Bay ${i + 1}`}
                          className="input-field"
                        />
                      ))}
                    </div>
                    <motion.button
                      whileHover={{ scale: 1.02 }}
                      whileTap={{ scale: 0.98 }}
                      onClick={autoFillHeights}
                      className="btn btn-primary"
                      style={{ width: '100%' }}
                    >
                      Auto-Fill Remaining Heights
                    </motion.button>
                  </motion.div>
                )}
              </motion.div>
            )}
          </AnimatePresence>

          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.3 }}
            style={{ display: 'flex', gap: '12px', marginTop: '28px' }}
          >
            <motion.button
              whileHover={{ scale: 1.02 }}
              whileTap={{ scale: 0.98 }}
              onClick={handleSaveElevation}
              className="btn btn-primary"
              disabled={isSaving}
              style={{ 
                flex: 1, 
                height: '50px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '8px',
                opacity: isSaving ? 0.7 : 1
              }}
            >
              {isSaving ? (
                <>
                  <motion.div
                    animate={{ rotate: 360 }}
                    transition={{ duration: 1, repeat: Infinity, ease: 'linear' }}
                    style={{
                      width: '16px',
                      height: '16px',
                      border: '2px solid rgba(255,255,255,0.3)',
                      borderTop: '2px solid white',
                      borderRadius: '50%'
                    }}
                  />
                  Saving...
                </>
              ) : (
                <>
                  <FiSave />
                  {isUpdateMode ? 'UPDATE ELEVATION' : 'CREATE ELEVATION'}
                </>
              )}
            </motion.button>
            <motion.button
              whileHover={{ scale: 1.1 }}
              whileTap={{ scale: 0.9 }}
              onClick={handleDeleteElevation}
              className="btn btn-danger"
              style={{ padding: '12px', minWidth: '50px' }}
            >
              <FiTrash2 />
            </motion.button>
          </motion.div>
        </motion.div>

        {/* Right Column: Door Manager */}
        <motion.div
          variants={slideIn}
          initial="hidden"
          animate="visible"
          transition={{ delay: 0.2 }}
          style={{ 
            flex: 1, 
            minWidth: '400px',
            background: 'linear-gradient(135deg, #1A1A1A 0%, #222222 100%)',
            borderRadius: '16px', 
            padding: '28px',
            boxShadow: '0 8px 24px rgba(0, 115, 230, 0.2)',
            border: '1px solid rgba(0, 115, 230, 0.2)'
          }}
        >
          <motion.h3
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.3 }}
            style={{ 
              color: COLOR_ACCENT, 
              marginBottom: '20px',
              fontSize: '18px',
              fontWeight: 'bold',
              display: 'flex',
              alignItems: 'center',
              gap: '8px'
            }}
          >
            <FiSettings />
            DOOR MANAGER
          </motion.h3>

          <motion.div
            variants={fadeIn}
            initial="hidden"
            animate="visible"
            style={{ marginBottom: '16px' }}
          >
            <label className="input-label">Size</label>
            <select 
              value={doorSize} 
              onChange={(e) => setDoorSize(e.target.value)} 
              className="input-field"
            >
              {doorOptions.filter(o => o !== 'None').map(opt => (
                <option key={opt} value={opt}>{opt}</option>
              ))}
            </select>
          </motion.div>

          <motion.div
            variants={fadeIn}
            initial="hidden"
            animate="visible"
            transition={{ delay: 0.1 }}
            style={{ marginBottom: '16px' }}
          >
            <label className="input-label">Count (Per Elevation)</label>
            <input 
              type="number" 
              value={doorCount} 
              onChange={(e) => setDoorCount(e.target.value)} 
              className="input-field" 
            />
          </motion.div>

          <motion.div
            variants={fadeIn}
            initial="hidden"
            animate="visible"
            transition={{ delay: 0.2 }}
            style={{ marginBottom: '16px' }}
          >
            <label className="input-label">Style</label>
            <select 
              value={doorStile} 
              onChange={(e) => setDoorStile(e.target.value)} 
              className="input-field"
            >
              {stileOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
            </select>
          </motion.div>

          <motion.div
            variants={fadeIn}
            initial="hidden"
            animate="visible"
            transition={{ delay: 0.3 }}
            style={{ marginBottom: '20px' }}
          >
            <label className="input-label">Hardware:</label>
            <div style={{ 
              maxHeight: '150px', 
              overflowY: 'auto',
              padding: '8px',
              background: 'rgba(0, 0, 0, 0.3)',
              borderRadius: '8px',
              marginTop: '8px'
            }}>
              {hardwareOptions.map(hw => (
                <motion.label
                  key={hw}
                  whileHover={{ x: 4 }}
                  style={{ 
                    display: 'block', 
                    marginTop: '8px', 
                    color: COLOR_TEXT_DIM,
                    cursor: 'pointer',
                    padding: '4px 8px',
                    borderRadius: '4px',
                    transition: 'background 0.2s'
                  }}
                  onMouseEnter={(e) => e.currentTarget.style.background = 'rgba(0, 115, 230, 0.1)'}
                  onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
                >
                  <input
                    type="checkbox"
                    checked={doorHardware[hw] || false}
                    onChange={(e) => setDoorHardware({ ...doorHardware, [hw]: e.target.checked })}
                    style={{ marginRight: '8px', cursor: 'pointer' }}
                  />
                  {hw}
                </motion.label>
              ))}
            </div>
          </motion.div>

          <motion.div
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.4 }}
            style={{ display: 'flex', gap: '10px', marginBottom: '24px' }}
          >
            <motion.button
              whileHover={{ scale: 1.02 }}
              whileTap={{ scale: 0.98 }}
              onClick={handleAddDoor}
              className="btn btn-primary"
              style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}
            >
              <FiPlus />
              ADD
            </motion.button>
            <motion.button
              whileHover={{ scale: 1.02 }}
              whileTap={{ scale: 0.98 }}
              onClick={handleUpdateDoor}
              className="btn btn-primary"
              style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}
            >
              <FiCheck />
              UPDATE
            </motion.button>
          </motion.div>

          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.5 }}
            style={{ 
              borderTop: `1px solid rgba(0, 115, 230, 0.2)`, 
              paddingTop: '20px', 
              maxHeight: '400px', 
              overflowY: 'auto' 
            }}
          >
            <AnimatePresence mode="popLayout">
              {doors.map((door, index) => {
                const hwTxt = Object.entries(door.hardware).filter(([_, v]) => v).map(([k]) => k).join(', ');
                return (
                  <motion.div
                    key={index}
                    layout
                    initial={{ opacity: 0, scale: 0.9, y: 20 }}
                    animate={{ opacity: 1, scale: 1, y: 0 }}
                    exit={{ opacity: 0, scale: 0.9, y: -20 }}
                    transition={{ type: 'spring', stiffness: 300, damping: 25 }}
                    whileHover={{ scale: 1.02, x: 4 }}
                    style={{ 
                      background: 'linear-gradient(135deg, #2A2A2A 0%, #333333 100%)',
                      padding: '16px', 
                      borderRadius: '12px', 
                      marginBottom: '12px',
                      border: '1px solid rgba(0, 115, 230, 0.2)',
                      boxShadow: '0 2px 8px rgba(0, 0, 0, 0.3)'
                    }}
                  >
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'start' }}>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontWeight: 'bold', color: COLOR_TEXT, marginBottom: '6px' }}>
                          Door {index + 1}
                        </div>
                        <div style={{ fontSize: '12px', color: COLOR_TEXT_DIM, marginBottom: '4px' }}>
                          {door.size} | {door.stile} Stile | Qty: {door.count}
                        </div>
                        {hwTxt && (
                          <div style={{ fontSize: '10px', color: COLOR_TEXT_DIM, fontStyle: 'italic' }}>
                            HW: {hwTxt}
                          </div>
                        )}
                      </div>
                      <div style={{ display: 'flex', gap: '4px' }}>
                        <motion.button
                          whileHover={{ scale: 1.2, rotate: 15 }}
                          whileTap={{ scale: 0.9 }}
                          onClick={() => handleEditDoor(index)}
                          className="icon-btn"
                          style={{ color: COLOR_ACCENT }}
                        >
                          <FiEdit />
                        </motion.button>
                        <motion.button
                          whileHover={{ scale: 1.2, rotate: -15 }}
                          whileTap={{ scale: 0.9 }}
                          onClick={() => handleDeleteDoor(index)}
                          className="icon-btn"
                          style={{ color: '#dc3545' }}
                        >
                          <FiTrash2 />
                        </motion.button>
                      </div>
                    </div>
                  </motion.div>
                );
              })}
            </AnimatePresence>
            {doors.length === 0 && (
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                style={{
                  textAlign: 'center',
                  padding: '40px 20px',
                  color: COLOR_TEXT_DIM
                }}
              >
                <FiSettings style={{ fontSize: '48px', marginBottom: '12px', opacity: 0.5 }} />
                <p>No doors added yet</p>
              </motion.div>
            )}
          </motion.div>
        </motion.div>
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
