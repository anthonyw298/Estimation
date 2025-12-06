// Simplified pricing utilities - core functionality
import { partsData } from '../data/partsData';
import { PART_NUMBER_MAP } from '../data/partNumber';

const EPSILON = 1e-9;

export const HORIZONTAL_BAY_PARTS = new Set(["BE9-2514", "BE9-2515", "E9-2519"]);
export const VERTICAL_BAY_PARTS = new Set(["E9-2512", "BE9-2511"]);
export const BAY_WIDTH_PARTS = new Set([...HORIZONTAL_BAY_PARTS, ...VERTICAL_BAY_PARTS]);

export function isBayWidthPart(partNumber: string, requestedQty?: number | number[], description?: string): boolean {
  if (partNumber === "BE9-2513") {
    if (requestedQty !== undefined) {
      return Array.isArray(requestedQty);
    }
    return description ? /sill|jamb|vertical/i.test(description) : false;
  }
  if (VERTICAL_BAY_PARTS.has(partNumber)) {
    return requestedQty === undefined || Array.isArray(requestedQty);
  }
  return HORIZONTAL_BAY_PARTS.has(partNumber);
}

export function parseLengthToFeet(lengthStr: string | null | undefined): number {
  if (!lengthStr || typeof lengthStr !== 'string' || !lengthStr.trim()) return 0.0;
  const normalized = lengthStr.replace(/'/g, "'").replace(/"/g, '"');
  let feet = 0.0;
  let inches = 0.0;

  // Search for feet (e.g., "10ft", "10'")
  const feetMatch = normalized.match(/(\d+\.?\d*)\s*(ft|')/i);
  if (feetMatch) feet = parseFloat(feetMatch[1]);

  // Search for inches - handle both decimal and fractional formats
  // First try to match fractional inches like "4-5/8"", "4 5/8"", etc.
  const fractionalInchMatch = normalized.match(/(\d+)\s*[- ]\s*(\d+)\s*\/\s*(\d+)\s*[""]/i);
  if (fractionalInchMatch) {
    const wholeInches = parseFloat(fractionalInchMatch[1]);
    const numerator = parseFloat(fractionalInchMatch[2]);
    const denominator = parseFloat(fractionalInchMatch[3]);
    inches = wholeInches + (numerator / denominator);
  } else {
    // Try to match decimal inches (e.g., "6in", "6\"", "6.5\"")
    const inchesMatch = normalized.match(/(\d+\.?\d*)\s*(in|")/i);
    if (inchesMatch) inches = parseFloat(inchesMatch[1]);
  }

  if (feet || inches) return feet + (inches / 12);

  // Fallback: if no units, assume the number itself is in feet
  const numMatch = normalized.match(/(\d+\.?\d*)/);
  if (numMatch) return parseFloat(numMatch[1]);

  return 0.0;
}

export interface ExtraMaterials {
  [key: string]: {
    quantity: number;
    length_pieces: number[];
  };
}

export function loadExtraMaterials(projectName: string): ExtraMaterials {
  try {
    const key = `ug_project_${projectName}_ExtraMaterials`;
    const data = localStorage.getItem(key);
    return data ? JSON.parse(data) : {};
  } catch {
    return {};
  }
}

export function saveExtraMaterials(projectName: string, materials: ExtraMaterials): void {
  try {
    const key = `ug_project_${projectName}_ExtraMaterials`;
    localStorage.setItem(key, JSON.stringify(materials));
  } catch (e) {
    console.error('Error saving extra materials:', e);
  }
}

export function getUnitPriceByPart(
  partNumber: string,
  finish?: string,
  projectName?: string
): [number | null, string] {
  const match = (partsData as any)[partNumber];
  if (!match) return [null, 'pcs'];

  const listPriceRaw = match['List Price'] || 0;
  const unitsStr = match['Units'] || null;
  const lengthStr = match['Length'] || null;

  let unitCount = 1;
  if (typeof unitsStr === 'string') {
    const unitsLower = unitsStr.toLowerCase().trim();
    if (unitsLower.includes('pcs') || unitsLower.includes('pc')) {
      const numPart = unitsLower.split('pc')[0].trim();
      if (numPart) {
        unitCount = parseInt(numPart) || 1;
      }
    }
  }

  let listPriceEffective = 0.0;
  let unitType = 'pcs';

  const isProfile = partNumber in (PART_NUMBER_MAP.profiles || {});
  // Gaskets should be treated as profiles for pricing (sold by length, with leftover tracking)
  const isGasket = ['E2-0052', 'E2-0053', 'E2-0065'].includes(partNumber);
  const treatAsProfile = isProfile || isGasket;

  if (treatAsProfile) {
    unitType = 'ft';
    // Gaskets are treated as profiles: sold by length, divide by length to get per-foot price
    if (Array.isArray(listPriceRaw) && listPriceRaw.length === 3) {
      const finishNorm = finish?.toLowerCase() || 'clear';
      if (finishNorm === 'clear') listPriceEffective = parseFloat(listPriceRaw[0]) || 0;
      else if (finishNorm === 'black') listPriceEffective = parseFloat(listPriceRaw[1]) || 0;
      else if (finishNorm === 'paint') listPriceEffective = parseFloat(listPriceRaw[2]) || 0;
      else listPriceEffective = parseFloat(listPriceRaw[0]) || 0;
    } else {
      listPriceEffective = parseFloat(listPriceRaw) || 0;
      if (!Array.isArray(listPriceRaw)) {
        const finishNorm = finish?.toLowerCase() || 'clear';
        if (finishNorm === 'black') listPriceEffective *= 1.1;
        else if (finishNorm === 'paint') listPriceEffective *= 1.2;
      }
    }

    const lengthFt = parseLengthToFeet(lengthStr);
    if (lengthFt > EPSILON) {
      listPriceEffective /= lengthFt;
    }
  } else {
    // For accessories and non-profiles, the length field is just a physical dimension, not a pricing unit
    // We should NOT divide by length for accessories - they are sold by pieces, not by length
    listPriceEffective = parseFloat(listPriceRaw) || 0;
    if (unitCount > 1) {
      listPriceEffective /= unitCount;
    }
    // Note: We intentionally do NOT divide by length for accessories
    // The length field in the database is just a physical dimension, not used for pricing
  }

  return [listPriceEffective, unitType];
}

export interface MaterialImpact {
  part_number: string;
  requested_qty: number | number[];
  purchased_qty_or_length: number;
  leftover_generated_qty_or_length: number;
  used_from_leftover_qty_or_length: number;
  cost_incurred: number;
  type_processed_as: string;
  finish?: string;
  description?: string;
}

export function getPriceByPart(
  partNumber: string,
  requestedQty: number | number[],
  finish?: string,
  currentExtraMaterials?: ExtraMaterials,
  summary: boolean = false,
  group: boolean = false,
  projectName?: string,
  description?: string
): [number | null, string, MaterialImpact | null] {
  const match = (partsData as any)[partNumber];
  if (!match) {
    const unitType = typeof requestedQty === 'number' && requestedQty % 1 !== 0 ? 'ft' : 'pcs';
    return [null, unitType, null];
  }

  const [unitPrice, unitType] = getUnitPriceByPart(partNumber, finish, projectName);
  if (unitPrice === null) {
    return [null, unitType, null];
  }

  const extraMaterialsKey = partNumber;
  // Gaskets are always treated as profiles for inventory purposes
  const isGasketForInventory = ['E2-0052', 'E2-0053', 'E2-0065'].includes(partNumber);
  const isProfileForInventory = partNumber in (PART_NUMBER_MAP.profiles || {}) || group || isGasketForInventory;
  const finalKey = isProfileForInventory && finish && finish.trim() !== '' ? `${partNumber}-${finish.toLowerCase()}` : partNumber;

  let partExtra: { quantity: number; length_pieces: number[] } = { quantity: 0, length_pieces: [] };
  if (!summary && projectName) {
    const extraMaterials = currentExtraMaterials || loadExtraMaterials(projectName);
    partExtra = extraMaterials[finalKey] || { quantity: 0, length_pieces: [] };
  }

  const materialImpact: MaterialImpact = {
    part_number: partNumber,
    requested_qty: requestedQty,
    purchased_qty_or_length: 0.0,
    leftover_generated_qty_or_length: 0.0,
    used_from_leftover_qty_or_length: 0.0,
    cost_incurred: 0.0,
    type_processed_as: '',
    finish: finish,
    description: description
  };

  let totalPrice = 0.0;

  if (isProfileForInventory) {
    materialImpact.type_processed_as = 'profile';
    const unitsStr = match['Units'] || '1 pcs.';
    const lengthStr = match['Length'] || '';
    const minPurchaseLength = parseLengthToFeet(lengthStr) || 1.0;
    const leftoverPieces = [...(partExtra.length_pieces || [])].sort((a, b) => b - a);

    if (Array.isArray(requestedQty) && isBayWidthPart(partNumber, requestedQty, description) && !summary) {
      // Bay width optimization logic - matches Python version
      const bayWidths = requestedQty.map(q => parseFloat(q.toString()));
      const totalNeeded = bayWidths.reduce((sum, w) => sum + w, 0);
      let usedFromLeftover = 0.0;
      const remainingLeftovers = [...leftoverPieces].sort((a, b) => b - a);
      const leftoverPiecesConsumed: Array<{ original_length: number; used_length: number }> = [];
      
      // Sort bay widths in descending order to match largest pieces to largest requirements first
      const sortedBayWidths = [...bayWidths].sort((a, b) => b - a);
      
      for (const bayWidth of sortedBayWidths) {
        let bestMatchIndex: number | null = null;
        let bestMatchLength: number | null = null;
        
        // Find the smallest leftover piece that is >= bay_width (closest fit to minimize waste)
        for (let i = 0; i < remainingLeftovers.length; i++) {
          const leftover = remainingLeftovers[i];
          if (leftover >= bayWidth - EPSILON) {
            if (bestMatchIndex === null || leftover < bestMatchLength!) {
              bestMatchIndex = i;
              bestMatchLength = leftover;
            }
          }
        }
        
        if (bestMatchIndex !== null) {
          // Use the closest fitting piece
          const leftover = remainingLeftovers[bestMatchIndex];
          usedFromLeftover += bayWidth;
          const remainingAfterUse = leftover - bayWidth;
          
          // Track this consumption
          leftoverPiecesConsumed.push({
            original_length: leftover,
            used_length: bayWidth
          });
          
          remainingLeftovers.splice(bestMatchIndex, 1);
          if (remainingAfterUse > EPSILON) {
            remainingLeftovers.push(remainingAfterUse);
            remainingLeftovers.sort((a, b) => b - a);
          }
        }
      }

      const remainingNeeded = totalNeeded - usedFromLeftover;
      if (remainingNeeded > EPSILON) {
        const numBundles = Math.ceil(remainingNeeded / minPurchaseLength);
        const actualPurchased = numBundles * minPurchaseLength;
        totalPrice = unitPrice * actualPurchased;
        const leftoverAfterUse = actualPurchased - remainingNeeded;
        materialImpact.purchased_qty_or_length = actualPurchased;
        materialImpact.leftover_generated_qty_or_length = leftoverAfterUse > EPSILON ? leftoverAfterUse : 0.0;
        materialImpact.cost_incurred = totalPrice;
      } else {
        totalPrice = 0.0;
        materialImpact.purchased_qty_or_length = 0.0;
        materialImpact.cost_incurred = 0.0;
      }
      materialImpact.used_from_leftover_qty_or_length = usedFromLeftover;
      
      // Add bay width list tracking fields
      (materialImpact as any).is_bay_width_list = true;
      (materialImpact as any).leftover_pieces_consumed = leftoverPiecesConsumed;
      (materialImpact as any).bay_widths_processed = bayWidths;
    } else {
      const qty = Array.isArray(requestedQty) ? requestedQty.reduce((sum, q) => sum + parseFloat(q.toString()), 0) : requestedQty;
      const suitableIndex = leftoverPieces.findIndex(p => p >= qty - EPSILON);

      if (suitableIndex >= 0 && !summary) {
        totalPrice = 0.0;
        materialImpact.used_from_leftover_qty_or_length = qty;
      } else {
        const numBundles = Math.ceil(qty / minPurchaseLength);
        const actualPurchased = numBundles * minPurchaseLength;
        totalPrice = unitPrice * actualPurchased;
        const leftoverPiece = Math.max(0.0, actualPurchased - qty);
        materialImpact.purchased_qty_or_length = actualPurchased;
        materialImpact.leftover_generated_qty_or_length = leftoverPiece > EPSILON ? leftoverPiece : 0.0;
        materialImpact.cost_incurred = totalPrice;
      }
    }
  } else {
    materialImpact.type_processed_as = 'accessory';
    const unitsStr = match['Units'] || '1 pcs.';
    let unitCountPerBundle = 1;
    if (typeof unitsStr === 'string' && unitsStr.toLowerCase().includes('pc')) {
      const numPart = unitsStr.toLowerCase().split('pc')[0].trim();
      if (numPart) unitCountPerBundle = parseInt(numPart) || 1;
    }

    const leftoverQty = partExtra.quantity || 0;
    const qty = Array.isArray(requestedQty) ? requestedQty.reduce((sum, q) => sum + parseFloat(q.toString()), 0) : requestedQty;
    const usedFromExisting = summary ? 0 : Math.min(qty, leftoverQty);
    const remainingNeeded = qty - usedFromExisting;
    
    // Initialize variables to match Python version
    let actualPurchased = 0;
    let excessQty = 0;
    totalPrice = 0.0;

    if (remainingNeeded > 0) {
      const numBundles = Math.ceil(remainingNeeded / unitCountPerBundle);
      actualPurchased = numBundles * unitCountPerBundle;
      // unitPrice is already per-piece (after dividing by unitCount in getUnitPriceByPart)
      // So multiply by actualPurchased (number of pieces) to get total price
      totalPrice = unitPrice * actualPurchased;
      excessQty = actualPurchased - remainingNeeded;
    }
    
    materialImpact.used_from_leftover_qty_or_length = usedFromExisting;
    materialImpact.purchased_qty_or_length = actualPurchased;
    materialImpact.leftover_generated_qty_or_length = excessQty;
    materialImpact.cost_incurred = totalPrice;
  }

  return summary ? [totalPrice, unitType, null] : [totalPrice, unitType, materialImpact];
}

export function getMultiplier(runningGrandTotal: number): number {
  return runningGrandTotal < 50000 ? 0.614 : 0.572;
}

export function applyMaterialImpactToExtraMaterialsInMemory(
  materialsDict: ExtraMaterials,
  materialImpact: MaterialImpact
): void {
  if (!materialImpact || materialImpact.part_number === 'N/A - Manual') return;

  const partNumber = materialImpact.part_number;
  const typeProcessedAs = materialImpact.type_processed_as;
  const finish = materialImpact.finish;
  if (!partNumber) return;

  // Construct the key for extra materials based on part number and finish (for profiles and gaskets)
  // Gaskets are treated as profiles for inventory purposes, so they also use finish in the key
  const isGasketForInventory = ['E2-0052', 'E2-0053', 'E2-0065'].includes(partNumber);
  const isProfileType = typeProcessedAs === 'profile' || isGasketForInventory;
  let extraMaterialsKey = partNumber;
  if (isProfileType && finish && finish.trim() !== '') {
    extraMaterialsKey = `${partNumber}-${finish.toLowerCase()}`;
  }

  if (!materialsDict[extraMaterialsKey]) {
    materialsDict[extraMaterialsKey] = { quantity: 0, length_pieces: [] };
  }
  const partExtra = materialsDict[extraMaterialsKey];

  // Use the isProfileType already declared above
  if (isProfileType) {
    // Handle bay width lists specially
    const isBayWidthList = (materialImpact as any).is_bay_width_list || false;
    const leftoverPiecesConsumed = (materialImpact as any).leftover_pieces_consumed || [];

    if (isBayWidthList && leftoverPiecesConsumed.length > 0) {
      // Process each consumed leftover piece
      const tempLeftovers = [...(partExtra.length_pieces || [])].sort((a, b) => b - a);
      for (const consumedInfo of leftoverPiecesConsumed) {
        const originalLength = consumedInfo.original_length;
        const usedLength = consumedInfo.used_length;

        // Find and remove the matching leftover piece
        let found = false;
        for (let i = 0; i < tempLeftovers.length; i++) {
          const pieceLen = tempLeftovers[i];
          // Match by original length (with tolerance)
          if (Math.abs(pieceLen - originalLength) < EPSILON) {
            const remainingAfterUse = pieceLen - usedLength;
            tempLeftovers.splice(i, 1);
            if (remainingAfterUse > EPSILON) {
              tempLeftovers.push(remainingAfterUse);
              tempLeftovers.sort((a, b) => b - a);
            }
            found = true;
            break;
          }
        }

        if (!found) {
          // Fallback: find any piece >= used_length
          for (let i = 0; i < tempLeftovers.length; i++) {
            const pieceLen = tempLeftovers[i];
            if (pieceLen >= usedLength - EPSILON) {
              const remainingAfterUse = pieceLen - usedLength;
              tempLeftovers.splice(i, 1);
              if (remainingAfterUse > EPSILON) {
                tempLeftovers.push(remainingAfterUse);
                tempLeftovers.sort((a, b) => b - a);
              }
              found = true;
              break;
            }
          }
        }
      }

      partExtra.length_pieces = tempLeftovers;
      partExtra.quantity = 0.0;
    } else {
      // Standard handling for single quantity
      const usedFromLeftover = materialImpact.used_from_leftover_qty_or_length || 0.0;
      if (usedFromLeftover > EPSILON) {
        const tempLeftovers = [...(partExtra.length_pieces || [])].sort((a, b) => b - a);
        let consumed = false;
        for (let i = 0; i < tempLeftovers.length; i++) {
          const pieceLen = tempLeftovers[i];
          if (pieceLen >= usedFromLeftover - EPSILON) {
            const remainingAfterUse = pieceLen - usedFromLeftover;
            tempLeftovers.splice(i, 1);
            if (remainingAfterUse > EPSILON) {
              tempLeftovers.push(remainingAfterUse);
            }
            consumed = true;
            break;
          }
        }
        partExtra.length_pieces = tempLeftovers;
      }
      partExtra.quantity = 0.0;
    }

    // Add leftover generated
    const leftoverGenerated = materialImpact.leftover_generated_qty_or_length || 0.0;
    if (leftoverGenerated > EPSILON) {
      partExtra.length_pieces.push(leftoverGenerated);
      partExtra.length_pieces.sort((a, b) => b - a);
      // Debug logging for gaskets
      if (['E2-0052', 'E2-0053', 'E2-0065'].includes(partNumber)) {
        console.log(`[Gasket Residual] Part: ${partNumber}, Key: ${extraMaterialsKey}, Leftover: ${leftoverGenerated}, All pieces:`, partExtra.length_pieces);
      }
    }
  } else if (typeProcessedAs === 'accessory') {
    const currentQty = partExtra.quantity || 0;
    const netChange = (materialImpact.leftover_generated_qty_or_length || 0.0) - 
                      (materialImpact.used_from_leftover_qty_or_length || 0.0);
    partExtra.quantity = Math.max(0, Math.round((currentQty + netChange) * 10000) / 10000);
    partExtra.length_pieces = [];
  }

  materialsDict[extraMaterialsKey] = partExtra;
}

