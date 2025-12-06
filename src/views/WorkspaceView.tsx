import { useState, useEffect } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
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

export default function WorkspaceView() {
  const location = useLocation();
  const navigate = useNavigate();
  const projectName = (location.state as any)?.projectName || '';

  const [elevations, setElevations] = useState<Record<string, ElevationData>>({});
  const [currentElevation, setCurrentElevation] = useState<string>('New Elevation');
  const [snackMessage, setSnackMessage] = useState('');

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

  const handleSaveElevation = () => {
    try {
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

  const isYes45 = system === 'YES 45TU FRONT SET(OG)';
  const elevationNames = ['New Elevation', ...Object.keys(elevations).sort()];

  return (
    <div style={{ backgroundColor: COLOR_BG, minHeight: '100vh', padding: '20px' }}>
      {/* Header */}
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: '20px', height: '60px' }}>
        <button onClick={() => navigate('/')} style={{ background: 'none', border: 'none', color: COLOR_TEXT, cursor: 'pointer', fontSize: '24px', marginRight: '10px' }}>
          ←
        </button>
        <h2 style={{ color: COLOR_TEXT, margin: 0 }}>{projectName.toUpperCase()}</h2>
        <div style={{ flex: 1 }} />
        <button 
          className="btn btn-primary"
          type="button"
          onClick={(e) => {
            e.preventDefault();
            e.stopPropagation();
            console.log('=== GENERATE REPORT BUTTON CLICKED ===');
            console.log('Project name:', projectName);
            console.log('Current elevations state:', elevations);
            
            (async () => {
              try {
                if (!projectName) {
                  console.warn('No project name');
                  showSnack('No project selected');
                  return;
                }
                
                // Reload elevations from storage to ensure we have the latest data
                const latestElevations = loadElevations(projectName);
                console.log('Latest elevations from storage:', latestElevations);
                
                if (!latestElevations || Object.keys(latestElevations).length === 0) {
                  console.warn('No elevations found');
                  showSnack('No elevations to generate report. Please create an elevation first.');
                  return;
                }
                
                console.log('Calling generateExcelReport...');
                showSnack('Generating report...');
                console.log('Starting report generation with elevations:', Object.keys(latestElevations));
                
                await generateExcelReport(projectName, latestElevations);
                
                console.log('Report generation completed successfully');
                showSnack('Report generated successfully! Check your downloads folder.');
              } catch (e: any) {
                console.error('Report generation error:', e);
                console.error('Error stack:', e.stack);
                console.error('Error details:', JSON.stringify(e, Object.getOwnPropertyNames(e)));
                showSnack(`Error generating report: ${e.message || e.toString()}`);
              }
            })();
          }}
        >
          GENERATE REPORT
        </button>
      </div>

      <div style={{ display: 'flex', gap: '20px' }}>
        {/* Left Column: Elevation Form */}
        <div style={{ flex: 2, backgroundColor: COLOR_SURFACE, borderRadius: '10px', padding: '20px' }}>
          <select
            value={currentElevation}
            onChange={(e) => handleElevationLoad(e.target.value)}
            className="input-field"
            style={{ marginBottom: '20px' }}
          >
            {elevationNames.map(name => (
              <option key={name} value={name}>{name}</option>
            ))}
          </select>

          <div style={{ marginBottom: '15px' }}>
            <label className="input-label">System</label>
            <select value={system} onChange={(e) => setSystem(e.target.value)} className="input-field">
              {systemOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
            </select>
          </div>

          <div style={{ marginBottom: '15px' }}>
            <label className="input-label">Finish</label>
            <select value={finish} onChange={(e) => setFinish(e.target.value)} className="input-field">
              {finishOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
            </select>
          </div>

          <div style={{ display: 'flex', gap: '10px', marginBottom: '15px' }}>
            <div style={{ flex: 1 }}>
              <label className="input-label">Elevation Type (Name)</label>
              <input type="text" value={elevType} onChange={(e) => setElevType(e.target.value)} className="input-field" />
            </div>
            <div style={{ flex: 1 }}>
              <label className="input-label">Quantity</label>
              <input type="number" value={count} onChange={(e) => setCount(e.target.value)} className="input-field" />
            </div>
          </div>

          <div style={{ marginTop: '20px', marginBottom: '15px' }}>
            <h3 style={{ color: COLOR_TEXT_DIM, fontSize: '12px', marginBottom: '10px' }}>DIMENSIONS</h3>
            <div style={{ display: 'flex', gap: '10px' }}>
              <div style={{ flex: 1 }}>
                <label className="input-label">Opening Width (")</label>
                <input type="number" value={width} onChange={(e) => setWidth(e.target.value)} className="input-field" />
              </div>
              <div style={{ flex: 1 }}>
                <label className="input-label">Opening Height (")</label>
                <input type="number" value={height} onChange={(e) => setHeight(e.target.value)} className="input-field" />
              </div>
            </div>
          </div>

          {isYes45 && (
            <>
              <div style={{ marginTop: '20px', marginBottom: '15px' }}>
                <h3 style={{ color: COLOR_TEXT_DIM, fontSize: '12px', marginBottom: '10px' }}>BAY CONFIGURATION</h3>
                <div style={{ display: 'flex', gap: '10px' }}>
                  <div style={{ flex: 1 }}>
                    <label className="input-label">Bays Wide</label>
                    <input type="number" value={baysWide} onChange={(e) => setBaysWide(e.target.value)} className="input-field" />
                  </div>
                  <div style={{ flex: 1 }}>
                    <label className="input-label">Bays Tall</label>
                    <input type="number" value={baysTall} onChange={(e) => setBaysTall(e.target.value)} className="input-field" />
                  </div>
                </div>
              </div>

              {parseInt(baysWide) > 0 && (
                <div style={{ marginTop: '15px', marginBottom: '15px' }}>
                  <label className="input-label">Custom Bay Widths (leave blank to auto-fill)</label>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '10px', marginBottom: '10px' }}>
                    {customBayWidths.map((w, i) => (
                      <input
                        key={i}
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
                  <button onClick={autoFillWidths} className="btn btn-primary" style={{ width: '100%' }}>
                    Auto-Fill Remaining Widths
                  </button>
                </div>
              )}

              {parseInt(baysTall) > 0 && (
                <div style={{ marginTop: '15px', marginBottom: '15px' }}>
                  <label className="input-label">Custom Bay Heights (leave blank to auto-fill)</label>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '10px', marginBottom: '10px' }}>
                    {customBayHeights.map((h, i) => (
                      <input
                        key={i}
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
                  <button onClick={autoFillHeights} className="btn btn-primary" style={{ width: '100%' }}>
                    Auto-Fill Remaining Heights
                  </button>
                </div>
              )}
            </>
          )}

          <div style={{ display: 'flex', gap: '10px', marginTop: '20px' }}>
            <button onClick={handleSaveElevation} className="btn btn-primary" style={{ flex: 1, height: '50px' }}>
              CREATE ELEVATION
            </button>
            <button onClick={handleDeleteElevation} className="btn btn-danger" style={{ padding: '10px' }}>
              🗑️
            </button>
          </div>
        </div>

        {/* Right Column: Door Manager */}
        <div style={{ flex: 1, backgroundColor: COLOR_SURFACE, borderRadius: '10px', padding: '20px' }}>
          <h3 style={{ color: COLOR_ACCENT, marginBottom: '15px' }}>DOOR MANAGER</h3>

          <div style={{ marginBottom: '15px' }}>
            <label className="input-label">Size</label>
            <select value={doorSize} onChange={(e) => setDoorSize(e.target.value)} className="input-field">
              {doorOptions.filter(o => o !== 'None').map(opt => <option key={opt} value={opt}>{opt}</option>)}
            </select>
          </div>

          <div style={{ marginBottom: '15px' }}>
            <label className="input-label">Count (Per Elevation)</label>
            <input type="number" value={doorCount} onChange={(e) => setDoorCount(e.target.value)} className="input-field" />
          </div>

          <div style={{ marginBottom: '15px' }}>
            <label className="input-label">Style</label>
            <select value={doorStile} onChange={(e) => setDoorStile(e.target.value)} className="input-field">
              {stileOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
            </select>
          </div>

          <div style={{ marginBottom: '15px' }}>
            <label className="input-label">Hardware:</label>
            {hardwareOptions.map(hw => (
              <label key={hw} style={{ display: 'block', marginTop: '5px', color: COLOR_TEXT_DIM }}>
                <input
                  type="checkbox"
                  checked={doorHardware[hw] || false}
                  onChange={(e) => setDoorHardware({ ...doorHardware, [hw]: e.target.checked })}
                  style={{ marginRight: '5px' }}
                />
                {hw}
              </label>
            ))}
          </div>

          <div style={{ display: 'flex', gap: '10px', marginBottom: '20px' }}>
            <button onClick={handleAddDoor} className="btn btn-primary" style={{ flex: 1 }}>ADD</button>
            <button onClick={handleUpdateDoor} className="btn btn-primary" style={{ flex: 1 }}>UPDATE</button>
          </div>

          <div style={{ borderTop: `1px solid ${COLOR_SURFACE}`, paddingTop: '20px', maxHeight: '400px', overflowY: 'auto' }}>
            {doors.map((door, index) => {
              const hwTxt = Object.entries(door.hardware).filter(([_, v]) => v).map(([k]) => k).join(', ');
              return (
                <div key={index} style={{ backgroundColor: COLOR_INPUT_BG, padding: '10px', borderRadius: '5px', marginBottom: '5px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'start' }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontWeight: 'bold', color: COLOR_TEXT }}>Door {index + 1}</div>
                      <div style={{ fontSize: '12px', color: COLOR_TEXT_DIM }}>
                        {door.size} | {door.stile} Stile | Qty: {door.count}
                      </div>
                      {hwTxt && <div style={{ fontSize: '10px', color: COLOR_TEXT_DIM, fontStyle: 'italic' }}>HW: {hwTxt}</div>}
                    </div>
                    <div>
                      <button onClick={() => handleEditDoor(index)} style={{ background: 'none', border: 'none', color: 'blue', cursor: 'pointer', marginRight: '5px' }}>✏️</button>
                      <button onClick={() => handleDeleteDoor(index)} style={{ background: 'none', border: 'none', color: 'red', cursor: 'pointer' }}>🗑️</button>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {snackMessage && <div className="snackbar">{snackMessage}</div>}
    </div>
  );
}

