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

  const feetMatch = normalized.match(/(\d+\.?\d*)\s*(ft|')/i);
  if (feetMatch) feet = parseFloat(feetMatch[1]);

  const inchesMatch = normalized.match(/(\d+\.?\d*)\s*(in|")/i);
  if (inchesMatch) inches = parseFloat(inchesMatch[1]);

  if (feet || inches) return feet + (inches / 12);

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

  if (isProfile) {
    unitType = 'ft';
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
    listPriceEffective = parseFloat(listPriceRaw) || 0;
    if (unitCount > 1) {
      listPriceEffective /= unitCount;
    }
    const lengthFt = parseLengthToFeet(lengthStr);
    if (lengthFt > 1) {
      listPriceEffective /= lengthFt;
      unitType = 'ft';
    }
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
  const isProfileForInventory = partNumber in (PART_NUMBER_MAP.profiles || {}) || group;
  const finalKey = isProfileForInventory && finish ? `${partNumber}-${finish.toLowerCase()}` : partNumber;

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
      // Bay width optimization logic (simplified)
      const bayWidths = requestedQty.map(q => parseFloat(q.toString()));
      const totalNeeded = bayWidths.reduce((sum, w) => sum + w, 0);
      let usedFromLeftover = 0.0;

      // Simple matching logic
      for (const bayWidth of bayWidths) {
        const suitableIndex = leftoverPieces.findIndex(p => p >= bayWidth - EPSILON);
        if (suitableIndex >= 0) {
          usedFromLeftover += bayWidth;
          const leftover = leftoverPieces[suitableIndex];
          leftoverPieces.splice(suitableIndex, 1);
          const remaining = leftover - bayWidth;
          if (remaining > EPSILON) {
            leftoverPieces.push(remaining);
            leftoverPieces.sort((a, b) => b - a);
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

    if (remainingNeeded > 0) {
      const numBundles = Math.ceil(remainingNeeded / unitCountPerBundle);
      const actualPurchased = numBundles * unitCountPerBundle;
      totalPrice = unitPrice * actualPurchased;
      const excessQty = actualPurchased - remainingNeeded;
      materialImpact.purchased_qty_or_length = actualPurchased;
      materialImpact.leftover_generated_qty_or_length = excessQty;
      materialImpact.cost_incurred = totalPrice;
    }
    materialImpact.used_from_leftover_qty_or_length = usedFromExisting;
  }

  return summary ? [totalPrice, unitType, null] : [totalPrice, unitType, materialImpact];
}

export function getMultiplier(runningGrandTotal: number): number {
  return runningGrandTotal < 50000 ? 0.614 : 0.572;
}

