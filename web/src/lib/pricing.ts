import { partsData } from '@/data/parts-data';
import { PART_NUMBER_MAP } from '@/data/part-number';
import type { MaterialImpactDetails, ExtraMaterial } from '@/types';

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

export const EPSILON = 1e-9;

/** Horizontal parts: use custom bay widths */
export const HORIZONTAL_BAY_PARTS = new Set<string>([
  'BE9-2514',
  'BE9-2515',
  'E9-2519',
]);

/** Vertical parts: use height/2 split */
export const VERTICAL_BAY_PARTS = new Set<string>([
  'E9-2512',
  'BE9-2511',
]);

/** Union of horizontal + vertical bay parts */
export const BAY_WIDTH_PARTS = new Set<string>([
  ...HORIZONTAL_BAY_PARTS,
  ...VERTICAL_BAY_PARTS,
]);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Check if a part uses bay width/height lists for waste optimization.
 * For BE9-2513, both sill (horizontal) and jamb (vertical) use lists.
 * Vertical parts (2513 jamb, 2512, 2511) use height/2 split.
 * We can detect this by checking if requestedQty is a list or by description.
 */
export function _isBayWidthPart(
  partNumber: string,
  requestedQty?: number | number[] | null,
  description?: string | null,
): boolean {
  if (partNumber === 'BE9-2513') {
    // Both sill and jamb use lists now (sill uses bay widths, jamb uses height/2)
    if (requestedQty != null) {
      return Array.isArray(requestedQty);
    }
    // Fallback to description check if requestedQty not available
    if (description) {
      const d = description.toLowerCase();
      return d.includes('sill') || d.includes('jamb') || d.includes('vertical');
    }
    return false;
  }

  // Check if it's a vertical or horizontal bay part
  if (VERTICAL_BAY_PARTS.has(partNumber)) {
    // Vertical parts always use lists (height/2 split)
    return requestedQty == null || Array.isArray(requestedQty);
  }

  return HORIZONTAL_BAY_PARTS.has(partNumber);
}

/**
 * Converts various length formats to total feet.
 * Handles "24' - 0\"", "10ft", "6in", etc.
 */
export function parseLengthToFeet(lengthStr: string | null | undefined): number {
  if (typeof lengthStr !== 'string' || !lengthStr.trim()) return 0.0;

  // Normalize smart quotes
  let s = lengthStr
    .replace(/\u2018|\u2019/g, "'")
    .replace(/\u201C|\u201D/g, '"');

  let feet = 0.0;
  let inches = 0.0;

  // Search for feet (e.g., "10ft", "10'")
  const feetMatch = s.match(/(\d+\.?\d*)\s*(ft|')/i);
  if (feetMatch) feet = parseFloat(feetMatch[1]);

  // Search for inches (e.g., "6in", "6\"")
  const inchMatch = s.match(/(\d+\.?\d*)\s*(in|")/i);
  if (inchMatch) inches = parseFloat(inchMatch[1]);

  // If feet or inches were found, calculate total
  if (feet || inches) return feet + inches / 12;

  // Fallback: if no units, assume the number itself is in feet
  const numMatch = s.match(/(\d+\.?\d*)/);
  if (numMatch) return parseFloat(numMatch[1]);

  return 0.0;
}

// ---------------------------------------------------------------------------
// Unit price
// ---------------------------------------------------------------------------

/**
 * Retrieves the base list price per unit (foot for profiles/items with length,
 * piece for accessories) for a given part number from partsData, considering
 * finish for profiles.
 *
 * Returns [unitPrice, unitType] or [null, null] when not found.
 */
export function getUnitPriceByPart(
  partNumber: string,
  finish?: string | null,
): [number | null, string | null] {
  const match = partsData[partNumber];
  if (!match) return [null, null];

  const listPriceRaw = match['List Price']; // number | [number, number, number]
  const unitsStr = match['Units'] ?? null;
  const lengthStr = match['Length'] ?? null;

  let unitCount = 1;
  if (typeof unitsStr === 'string') {
    const unitsLower = unitsStr.toLowerCase().trim();
    if (unitsLower.includes('pc')) {
      const numPart = unitsLower.split('pc')[0].trim();
      if (numPart) {
        const parsed = parseInt(numPart, 10);
        if (!Number.isNaN(parsed)) unitCount = parsed;
      }
    }
  }

  let listPriceEffective = 0.0;
  let unitType = 'pcs'; // default

  const profileKeys = PART_NUMBER_MAP['profiles']
    ? Object.keys(PART_NUMBER_MAP['profiles'])
    : [];
  const isProfile = profileKeys.includes(partNumber);

  if (isProfile) {
    unitType = 'ft';

    // 1. Select base price based on finish (if list)
    if (Array.isArray(listPriceRaw) && listPriceRaw.length === 3) {
      const finishNorm = finish ? finish.toLowerCase() : 'clear';
      if (finishNorm === 'clear') {
        listPriceEffective = Number(listPriceRaw[0]);
      } else if (finishNorm === 'black') {
        listPriceEffective = Number(listPriceRaw[1]);
      } else if (finishNorm === 'paint') {
        listPriceEffective = Number(listPriceRaw[2]);
      } else {
        listPriceEffective = Number(listPriceRaw[0]);
      }
    } else {
      listPriceEffective = Number(listPriceRaw) || 0.0;
    }

    // 2. Apply the finish-based multiplier for profiles
    //    Only if List Price was NOT a 3-element array (already picked the correct price)
    if (!(Array.isArray(listPriceRaw) && listPriceRaw.length === 3)) {
      let finishMultiplier = 1.0;
      const finishNorm = finish ? finish.toLowerCase() : 'clear';
      if (finishNorm === 'black') finishMultiplier = 1.1;
      else if (finishNorm === 'paint') finishMultiplier = 1.2;
      listPriceEffective *= finishMultiplier;
    }

    // 3. Divide by length to get per-foot price for profiles
    const lengthFt = parseLengthToFeet(lengthStr);
    if (lengthFt > EPSILON) {
      listPriceEffective /= lengthFt;
    } else {
      console.warn(
        `Warning: Profile '${partNumber}' has zero or invalid length '${lengthStr}'. Unit price might be incorrect.`,
      );
    }
  } else {
    // Not a profile (accessory or other non-profile item)
    listPriceEffective = Number(listPriceRaw) || 0.0;

    // Apply unit_count division (e.g., if price is for a pack of 5, divide by 5)
    if (unitCount > 1) {
      listPriceEffective /= unitCount;
    }

    // Apply the length-based division if length is significant (> 1 foot)
    const lengthFt = parseLengthToFeet(lengthStr);
    if (lengthFt > 1) {
      listPriceEffective /= lengthFt;
      unitType = 'ft';
    }
  }

  return [listPriceEffective, unitType];
}

// ---------------------------------------------------------------------------
// Main pricing function
// ---------------------------------------------------------------------------

/**
 * Calculate price and material impact, considering finish for profiles.
 *
 * @param partNumber       The part number.
 * @param requestedQty     The quantity requested. Can be a list for bay width parts.
 * @param finish           The finish type ('clear', 'black', 'paint'). Only relevant for profiles.
 * @param currentExtraMaterials  In-memory extra materials state.
 * @param summary          If true, only return price and unit type, no material impact details.
 * @param group            If true, forces profile-like behavior for pricing.
 * @param description      Description of the part, used to distinguish jamb vs sill for BE9-2513.
 *
 * @returns [totalPrice, unitType, materialImpactDetails | null]
 */
export function getPriceByPart(
  partNumber: string,
  requestedQty: number | number[],
  finish?: string | null,
  currentExtraMaterials?: Record<string, ExtraMaterial> | null,
  summary = false,
  group = false,
  description?: string | null,
): [number | null, string, MaterialImpactDetails | null] {
  const match = partsData[partNumber];
  if (!match) {
    const fallbackUnit =
      typeof requestedQty === 'number' && requestedQty % 1 !== 0
        ? 'ft'
        : 'pcs';
    return [null, fallbackUnit, null];
  }

  const unitsStr = match['Units'] ?? '1 pcs.';
  const lengthStr = match['Length'] ?? '';

  // Use getUnitPriceByPart to get the correct unit price
  let [unitPrice, unitType] = getUnitPriceByPart(partNumber, finish);

  if (unitPrice == null) {
    return [null, unitType ?? 'pcs', null];
  }

  let totalPrice = 0.0;

  // Construct the key for extra materials based on part number and finish (for profiles)
  let extraMaterialsKey = partNumber;
  const profileKeys = PART_NUMBER_MAP['profiles']
    ? Object.keys(PART_NUMBER_MAP['profiles'])
    : [];
  const isProfileForInventory = profileKeys.includes(partNumber) || group;
  if (isProfileForInventory && finish) {
    extraMaterialsKey = `${partNumber}-${finish.toLowerCase()}`;
  }

  let partExtraSim: ExtraMaterial = { quantity: 0, length_pieces: [] };
  if (!summary) {
    if (currentExtraMaterials == null) {
      currentExtraMaterials = {};
    }
    const existing = currentExtraMaterials[extraMaterialsKey];
    if (existing) {
      partExtraSim = {
        quantity: existing.quantity ?? 0,
        length_pieces: [...(existing.length_pieces ?? [])],
      };
    }
  }

  const materialImpactDetails: MaterialImpactDetails = {
    part_number: partNumber,
    requested_qty: requestedQty,
    purchased_qty_or_length: 0.0,
    leftover_generated_qty_or_length: 0.0,
    used_from_leftover_qty_or_length: 0.0,
    cost_incurred: 0.0,
    type_processed_as: null,
    finish: finish ?? undefined,
    description: description ?? undefined,
  };

  if (isProfileForInventory) {
    // -----------------------------------------------------------------------
    // Profile / length-based logic
    // -----------------------------------------------------------------------
    unitType = 'ft';
    const minPurchaseLength = parseLengthToFeet(lengthStr) || 1.0;
    const leftoverPiecesSim = [...(partExtraSim.length_pieces ?? [])].sort(
      (a, b) => b - a,
    ); // Sort descending

    const partDescription = description ?? materialImpactDetails.description;
    const isBayWidthPart = _isBayWidthPart(
      partNumber,
      requestedQty,
      partDescription,
    );

    // ------------------------------------------------------------------
    // Bay-width list handling
    // ------------------------------------------------------------------
    if (
      Array.isArray(requestedQty) &&
      isBayWidthPart &&
      !summary
    ) {
      const bayWidths = requestedQty.map(Number);
      const totalNeeded = bayWidths.reduce((s, v) => s + v, 0);
      let usedFromLeftover = 0.0;
      const remainingLeftovers = [...leftoverPiecesSim];
      const matchedBays: number[] = [];

      // Track which leftover pieces were used
      const leftoverPiecesConsumed: Array<{
        original_length: number;
        used_length: number;
      }> = [];

      // Sort bay widths descending to match largest pieces to largest requirements first
      const sortedBayWidths = [...bayWidths].sort((a, b) => b - a);

      for (const bayWidth of sortedBayWidths) {
        let bestMatchIndex: number | null = null;
        let bestMatchLength: number | null = null;

        // Find the smallest leftover piece that is >= bayWidth (closest fit)
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
          const leftover = remainingLeftovers[bestMatchIndex];
          usedFromLeftover += bayWidth;
          const remainingAfterUse = leftover - bayWidth;
          leftoverPiecesConsumed.push({
            original_length: leftover,
            used_length: bayWidth,
          });
          remainingLeftovers.splice(bestMatchIndex, 1);
          if (remainingAfterUse > EPSILON) {
            remainingLeftovers.push(remainingAfterUse);
            remainingLeftovers.sort((a, b) => b - a);
          }
          matchedBays.push(bayWidth);
        } else {
          matchedBays.push(bayWidth);
        }
      }

      const remainingNeeded = totalNeeded - usedFromLeftover;

      if (remainingNeeded > EPSILON) {
        // Purchase new material for remaining needs
        const numBundlesNeeded = Math.ceil(remainingNeeded / minPurchaseLength);
        const actualPurchasedLength = numBundlesNeeded * minPurchaseLength;
        totalPrice = unitPrice * actualPurchasedLength;

        // Optimize cuts across new sticks
        const newLeftovers: number[] = [];
        // Sort bay widths descending for optimal packing
        const remainingBaysToCut = [...bayWidths].sort((a, b) => b - a);

        // Count pieces by size
        const pieceCounts = new Map<number, number>();
        for (const piece of remainingBaysToCut) {
          pieceCounts.set(piece, (pieceCounts.get(piece) ?? 0) + 1);
        }
        const uniqueSizes = [...pieceCounts.keys()].sort((a, b) => b - a);

        // Create a working copy of piece counts
        const remainingCounts = new Map<number, number>(pieceCounts);

        // Helper: total remaining pieces
        const totalRemainingPieces = () => {
          let sum = 0;
          for (const v of remainingCounts.values()) sum += v;
          return sum;
        };

        // Process sticks one at a time
        for (
          let stickNum = 0;
          stickNum < numBundlesNeeded;
          stickNum++
        ) {
          let currentStickRemaining = minPurchaseLength;
          const _piecesUsedThisStick: number[] = [];
          const isLastStick = stickNum === numBundlesNeeded - 1;

          // Greedy algorithm: repeatedly find pieces that fit
          while (currentStickRemaining > EPSILON) {
            let bestFitSize: number | null = null;
            let bestFitLength = 0.0;
            let isExactFit = false;
            let numPiecesToUse = 1;

            // First pass: look for exact fits (single piece that fills the stick completely)
            for (const size of uniqueSizes) {
              if ((remainingCounts.get(size) ?? 0) > 0) {
                if (Math.abs(currentStickRemaining - size) < EPSILON) {
                  bestFitSize = size;
                  bestFitLength = size;
                  isExactFit = true;
                  numPiecesToUse = 1;
                  break;
                }
              }
            }

            // Second pass: for earlier sticks, check if we can fill exactly with multiple smaller pieces
            if (!isExactFit && !isLastStick) {
              // Check smaller pieces first (reversed)
              const reversedSizes = [...uniqueSizes].reverse();
              for (const size of reversedSizes) {
                if (
                  (remainingCounts.get(size) ?? 0) > 0 &&
                  size <= currentStickRemaining + EPSILON
                ) {
                  const numFit = Math.floor(currentStickRemaining / size);
                  if (
                    numFit <= (remainingCounts.get(size) ?? 0) &&
                    Math.abs(currentStickRemaining - numFit * size) < EPSILON
                  ) {
                    bestFitSize = size;
                    bestFitLength = size;
                    isExactFit = true;
                    numPiecesToUse = numFit;
                    break;
                  }
                }
              }
            }

            // If no exact fit found, find best single piece that fits
            if (!isExactFit) {
              if (isLastStick) {
                // Last stick: prefer larger pieces
                for (const size of uniqueSizes) {
                  if (
                    (remainingCounts.get(size) ?? 0) > 0 &&
                    size <= currentStickRemaining + EPSILON
                  ) {
                    if (size > bestFitLength) {
                      bestFitSize = size;
                      bestFitLength = size;
                    }
                  }
                }
              } else {
                // Earlier sticks: prefer smaller pieces to fill completely
                const reversedSizes = [...uniqueSizes].reverse();
                for (const size of reversedSizes) {
                  if (
                    (remainingCounts.get(size) ?? 0) > 0 &&
                    size <= currentStickRemaining + EPSILON
                  ) {
                    if (bestFitSize === null || size > bestFitLength) {
                      bestFitSize = size;
                      bestFitLength = size;
                    }
                  }
                }
              }
            }

            if (bestFitSize !== null) {
              if (isExactFit && numPiecesToUse > 1) {
                for (let k = 0; k < numPiecesToUse; k++) {
                  _piecesUsedThisStick.push(bestFitSize);
                  currentStickRemaining -= bestFitSize;
                  remainingCounts.set(
                    bestFitSize,
                    (remainingCounts.get(bestFitSize) ?? 0) - 1,
                  );
                }
              } else {
                _piecesUsedThisStick.push(bestFitSize);
                currentStickRemaining -= bestFitSize;
                remainingCounts.set(
                  bestFitSize,
                  (remainingCounts.get(bestFitSize) ?? 0) - 1,
                );
              }
            } else {
              // No more pieces fit in this stick
              break;
            }
          }

          // Save leftover from this stick (only if less than full stick)
          if (
            currentStickRemaining > EPSILON &&
            currentStickRemaining < minPurchaseLength - EPSILON
          ) {
            newLeftovers.push(currentStickRemaining);
          }

          // If all pieces are used, no need for more sticks
          if (totalRemainingPieces() === 0) break;
        }

        materialImpactDetails.purchased_qty_or_length = actualPurchasedLength;
        materialImpactDetails.cost_incurred = totalPrice;
        // Store new leftovers - filter out zero, negative, or full-stick leftovers
        const validLeftovers = newLeftovers.filter(
          (lo) => lo > EPSILON && lo < minPurchaseLength - EPSILON,
        );
        if (validLeftovers.length > 0) {
          materialImpactDetails.all_new_leftovers = [...validLeftovers].sort(
            (a, b) => b - a,
          );
          materialImpactDetails.leftover_generated_qty_or_length = Math.max(
            ...validLeftovers,
          );
        } else {
          materialImpactDetails.all_new_leftovers = [];
          materialImpactDetails.leftover_generated_qty_or_length = 0.0;
        }
      } else {
        // All needs met from leftovers
        totalPrice = 0.0;
        materialImpactDetails.purchased_qty_or_length = 0.0;
        materialImpactDetails.cost_incurred = 0.0;
      }

      materialImpactDetails.used_from_leftover_qty_or_length = usedFromLeftover;
      materialImpactDetails.bay_widths_processed = matchedBays;
      materialImpactDetails.is_bay_width_list = true;
      materialImpactDetails.leftover_pieces_consumed = leftoverPiecesConsumed;

    } else if (Array.isArray(requestedQty) && !summary) {
      // ------------------------------------------------------------------
      // Standard list handling (non-bay-width parts or non-summary)
      // ------------------------------------------------------------------
      const piecesNeeded = [...requestedQty.map(Number)].sort((a, b) => b - a);
      const totalNeeded = piecesNeeded.reduce((s, v) => s + v, 0);

      // Track leftover pieces we'll use (for removal) and new leftovers we'll create
      const leftoverPiecesToUse: Array<[number, number]> = [];
      const remainingPieces = [...piecesNeeded];
      const tempLeftovers = !summary
        ? [...leftoverPiecesSim].sort((a, b) => b - a)
        : [];

      // First pass: try to use existing leftover pieces, matching closest fit
      for (const pieceLen of piecesNeeded) {
        let bestMatchIndex: number | null = null;
        let bestMatchLength: number | null = null;

        for (let i = 0; i < tempLeftovers.length; i++) {
          const leftoverLen = tempLeftovers[i];
          if (leftoverLen >= pieceLen - EPSILON) {
            if (bestMatchIndex === null || leftoverLen < bestMatchLength!) {
              bestMatchIndex = i;
              bestMatchLength = leftoverLen;
            }
          }
        }

        if (bestMatchIndex !== null) {
          const usedLeftover = tempLeftovers[bestMatchIndex];
          leftoverPiecesToUse.push([bestMatchIndex, pieceLen]);
          const remainingAfterUse = usedLeftover - pieceLen;
          tempLeftovers.splice(bestMatchIndex, 1);
          if (remainingAfterUse > EPSILON) {
            tempLeftovers.push(remainingAfterUse);
            tempLeftovers.sort((a, b) => b - a);
          }
          // Remove from remainingPieces (first occurrence)
          const idx = remainingPieces.indexOf(pieceLen);
          if (idx !== -1) remainingPieces.splice(idx, 1);
        }
      }

      // Calculate how much we still need to purchase
      const remainingNeeded = remainingPieces.length > 0
        ? remainingPieces.reduce((s, v) => s + v, 0)
        : 0.0;

      // Calculate total used from leftovers
      const usedFromLeftover = totalNeeded - remainingNeeded;

      if (remainingNeeded > EPSILON) {
        // Purchase new material for remaining pieces
        const numBundlesNeeded = Math.ceil(
          remainingNeeded / minPurchaseLength,
        );
        const actualPurchasedLength = numBundlesNeeded * minPurchaseLength;
        totalPrice = unitPrice * actualPurchasedLength;

        // Optimize cuts across new sticks
        const remainingPiecesSorted = [...remainingPieces].sort(
          (a, b) => b - a,
        );
        const newLeftovers: number[] = [];

        for (
          let stickNum = 0;
          stickNum < numBundlesNeeded;
          stickNum++
        ) {
          let currentStickRemaining = minPurchaseLength;
          const _piecesUsedThisStick: number[] = [];

          // Greedy algorithm: repeatedly find the largest piece that fits
          while (
            remainingPiecesSorted.length > 0 &&
            currentStickRemaining > EPSILON
          ) {
            let bestFitIndex: number | null = null;
            let bestFitLength = 0.0;

            for (let i = 0; i < remainingPiecesSorted.length; i++) {
              const pieceLen = remainingPiecesSorted[i];
              if (pieceLen <= currentStickRemaining + EPSILON) {
                if (pieceLen > bestFitLength) {
                  bestFitIndex = i;
                  bestFitLength = pieceLen;
                }
              }
            }

            if (bestFitIndex !== null) {
              const pieceLen = remainingPiecesSorted[bestFitIndex];
              currentStickRemaining -= pieceLen;
              _piecesUsedThisStick.push(pieceLen);
              remainingPiecesSorted.splice(bestFitIndex, 1);

              if (currentStickRemaining < EPSILON) break;
            } else {
              break;
            }
          }

          // Save leftover from this stick (only if it's less than a full stick)
          if (
            currentStickRemaining > EPSILON &&
            currentStickRemaining < minPurchaseLength - EPSILON
          ) {
            newLeftovers.push(currentStickRemaining);
          }

          // If we've used all pieces, no need for more sticks
          if (remainingPiecesSorted.length === 0) break;
        }

        // Sort leftovers (largest first)
        newLeftovers.sort((a, b) => b - a);

        materialImpactDetails.purchased_qty_or_length = actualPurchasedLength;
        materialImpactDetails.cost_incurred = totalPrice;
        const validLeftovers = newLeftovers.filter(
          (lo) => lo > EPSILON && lo < minPurchaseLength - EPSILON,
        );
        if (validLeftovers.length > 0) {
          materialImpactDetails.all_new_leftovers = [...validLeftovers].sort(
            (a, b) => b - a,
          );
          materialImpactDetails.leftover_generated_qty_or_length = Math.max(
            ...validLeftovers,
          );
        } else {
          materialImpactDetails.all_new_leftovers = [];
          materialImpactDetails.leftover_generated_qty_or_length = 0.0;
        }
      } else {
        // All needs met from leftovers
        totalPrice = 0.0;
        materialImpactDetails.purchased_qty_or_length = 0.0;
        materialImpactDetails.cost_incurred = 0.0;
      }

      materialImpactDetails.used_from_leftover_qty_or_length = usedFromLeftover;
      materialImpactDetails.leftover_pieces_consumed = leftoverPiecesToUse;

    } else {
      // ------------------------------------------------------------------
      // Single quantity – use existing logic
      // ------------------------------------------------------------------
      const qty = Array.isArray(requestedQty)
        ? requestedQty.reduce((s, v) => s + Number(v), 0)
        : Number(requestedQty);

      let suitableIndex: number | null = null;

      if (!summary) {
        // Find the closest-fitting leftover piece (smallest that fits)
        let bestFitIndex: number | null = null;
        let bestFitLength: number | null = null;
        for (let i = 0; i < leftoverPiecesSim.length; i++) {
          const pieceLen = leftoverPiecesSim[i];
          if (pieceLen >= qty - EPSILON) {
            if (bestFitIndex === null || pieceLen < bestFitLength!) {
              bestFitIndex = i;
              bestFitLength = pieceLen;
            }
          }
        }
        suitableIndex = bestFitIndex;
      }

      if (suitableIndex !== null) {
        // Material is taken from existing leftovers
        totalPrice = 0.0;
        materialImpactDetails.used_from_leftover_qty_or_length = qty;
      } else {
        // No suitable leftover found, purchase new material
        const numBundlesNeeded = Math.ceil(qty / minPurchaseLength);
        const actualPurchasedLength = numBundlesNeeded * minPurchaseLength;
        totalPrice = unitPrice * actualPurchasedLength;

        const leftoverPiece = Math.max(0.0, actualPurchasedLength - qty);

        materialImpactDetails.purchased_qty_or_length = actualPurchasedLength;
        materialImpactDetails.cost_incurred = totalPrice;
        if (
          leftoverPiece > EPSILON &&
          leftoverPiece < minPurchaseLength - EPSILON
        ) {
          materialImpactDetails.leftover_generated_qty_or_length =
            leftoverPiece;
          materialImpactDetails.all_new_leftovers = [leftoverPiece];
        } else {
          materialImpactDetails.leftover_generated_qty_or_length = 0.0;
          materialImpactDetails.all_new_leftovers = [];
        }
      }
    }

    materialImpactDetails.type_processed_as = 'profile';

  } else {
    // -------------------------------------------------------------------
    // Accessory / simple item (piece-based)
    // -------------------------------------------------------------------
    unitType = 'pcs';

    let unitCountPerBundle = 1;
    if (unitsStr.toLowerCase().includes('pc')) {
      const numPart = unitsStr.toLowerCase().split('pc')[0].trim();
      if (numPart) {
        const parsed = parseInt(numPart, 10);
        if (!Number.isNaN(parsed) && parsed > 0) unitCountPerBundle = parsed;
      }
    }

    const qty = Array.isArray(requestedQty)
      ? requestedQty.reduce((s, v) => s + Number(v), 0)
      : Number(requestedQty);

    const leftoverQtySim = partExtraSim.quantity ?? 0;
    let usedFromExistingLeftover = 0;
    if (!summary) usedFromExistingLeftover = Math.min(qty, leftoverQtySim);

    const remainingNeededQty = qty - usedFromExistingLeftover;
    let actualPurchasedQty = 0;
    let excessQtyFromNewPurchase = 0;
    totalPrice = 0.0;

    if (remainingNeededQty > 0) {
      const numBundlesNeeded = Math.ceil(
        remainingNeededQty / unitCountPerBundle,
      );
      actualPurchasedQty = numBundlesNeeded * unitCountPerBundle;
      totalPrice = unitPrice * actualPurchasedQty;
      excessQtyFromNewPurchase = actualPurchasedQty - remainingNeededQty;
    }

    materialImpactDetails.used_from_leftover_qty_or_length =
      usedFromExistingLeftover;
    materialImpactDetails.purchased_qty_or_length = actualPurchasedQty;
    materialImpactDetails.leftover_generated_qty_or_length =
      excessQtyFromNewPurchase;
    materialImpactDetails.cost_incurred = totalPrice;
    materialImpactDetails.type_processed_as = 'accessory';
  }

  if (summary) {
    return [totalPrice, unitType ?? 'pcs', null];
  }
  return [totalPrice, unitType ?? 'pcs', materialImpactDetails];
}

// ---------------------------------------------------------------------------
// Apply material impact – in memory
// ---------------------------------------------------------------------------

/**
 * Applies a single item's material impact to a provided materials dictionary
 * in memory. This is the TypeScript equivalent of
 * `apply_material_impact_to_extra_materials_in_memory`.
 */
export function applyMaterialImpactInMemory(
  materialsDict: Record<string, ExtraMaterial>,
  materialImpactDetails: MaterialImpactDetails | null | undefined,
): void {
  if (
    !materialImpactDetails ||
    materialImpactDetails.part_number === 'N/A - Manual'
  )
    return;

  const partNumber = materialImpactDetails.part_number;
  const typeProcessedAs = materialImpactDetails.type_processed_as;
  const finish = materialImpactDetails.finish;
  if (!partNumber) return;

  // Construct the key for extra materials based on part number and finish (for profiles)
  let extraMaterialsKey = partNumber;
  if (typeProcessedAs === 'profile' && finish) {
    extraMaterialsKey = `${partNumber}-${finish.toLowerCase()}`;
  }

  let partExtra: ExtraMaterial = materialsDict[extraMaterialsKey]
    ? {
        quantity: materialsDict[extraMaterialsKey].quantity ?? 0,
        length_pieces: [
          ...(materialsDict[extraMaterialsKey].length_pieces ?? []),
        ],
      }
    : { quantity: 0, length_pieces: [] };

  if (typeProcessedAs === 'profile') {
    // Handle bay width lists specially
    const isBayWidthList =
      materialImpactDetails.is_bay_width_list ?? false;
    const leftoverPiecesConsumed =
      materialImpactDetails.leftover_pieces_consumed ?? [];

    if (isBayWidthList && leftoverPiecesConsumed.length > 0) {
      // Process each consumed leftover piece (bay width format: dict with original_length and used_length)
      const tempLeftovers = [...(partExtra.length_pieces ?? [])].sort(
        (a, b) => b - a,
      );
      for (const consumedInfo of leftoverPiecesConsumed) {
        // Bay width format: { original_length, used_length }
        const originalLength = (consumedInfo as { original_length: number; used_length: number }).original_length;
        const usedLength = (consumedInfo as { original_length: number; used_length: number }).used_length;

        // Find and remove the matching leftover piece
        let found = false;
        for (let i = 0; i < tempLeftovers.length; i++) {
          if (Math.abs(tempLeftovers[i] - originalLength) < EPSILON) {
            const remainingAfterUse = tempLeftovers[i] - usedLength;
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
          // Fallback: find any piece >= usedLength
          for (let i = 0; i < tempLeftovers.length; i++) {
            if (tempLeftovers[i] >= usedLength - EPSILON) {
              const remainingAfterUse = tempLeftovers[i] - usedLength;
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
            console.warn(
              `[WARNING] (in-memory): Could not find suitable leftover piece to consume ${usedLength.toFixed(4)} for ${partNumber} (${finish}).`,
            );
          }
        }
      }

      partExtra.length_pieces = tempLeftovers;
      partExtra.quantity = 0.0;
    } else if (
      leftoverPiecesConsumed &&
      Array.isArray(leftoverPiecesConsumed) &&
      leftoverPiecesConsumed.length > 0
    ) {
      // Handle non-bay-width list format: tuples of [index, length_used]
      const tempLeftovers = [...(partExtra.length_pieces ?? [])].sort(
        (a, b) => b - a,
      );
      for (const consumedTuple of leftoverPiecesConsumed) {
        if (Array.isArray(consumedTuple) && consumedTuple.length === 2) {
          const [, lengthUsed] = consumedTuple as [number, number];
          // Find the closest-fitting leftover piece
          let bestMatchIndex: number | null = null;
          let bestMatchLength: number | null = null;
          for (let i = 0; i < tempLeftovers.length; i++) {
            if (tempLeftovers[i] >= lengthUsed - EPSILON) {
              if (
                bestMatchIndex === null ||
                tempLeftovers[i] < bestMatchLength!
              ) {
                bestMatchIndex = i;
                bestMatchLength = tempLeftovers[i];
              }
            }
          }

          if (bestMatchIndex !== null) {
            const usedPiece = tempLeftovers[bestMatchIndex];
            const remainingAfterUse = usedPiece - lengthUsed;
            tempLeftovers.splice(bestMatchIndex, 1);
            if (remainingAfterUse > EPSILON) {
              tempLeftovers.push(remainingAfterUse);
              tempLeftovers.sort((a, b) => b - a);
            }
          } else {
            console.warn(
              `[WARNING] (in-memory): Could not find suitable leftover piece to consume ${lengthUsed.toFixed(4)} for ${partNumber} (${finish}).`,
            );
          }
        }
      }

      partExtra.length_pieces = tempLeftovers;
      partExtra.quantity = 0.0;
    } else {
      // Standard handling for single quantity or list with no leftover consumption
      const usedFromLeftover =
        materialImpactDetails.used_from_leftover_qty_or_length ?? 0.0;
      if (usedFromLeftover > EPSILON) {
        const tempLeftovers = [...(partExtra.length_pieces ?? [])].sort(
          (a, b) => b - a,
        );
        // Find closest-fitting leftover piece
        let bestMatchIndex: number | null = null;
        let bestMatchLength: number | null = null;
        for (let i = 0; i < tempLeftovers.length; i++) {
          if (tempLeftovers[i] >= usedFromLeftover - EPSILON) {
            if (
              bestMatchIndex === null ||
              tempLeftovers[i] < bestMatchLength!
            ) {
              bestMatchIndex = i;
              bestMatchLength = tempLeftovers[i];
            }
          }
        }

        if (bestMatchIndex !== null) {
          const usedPiece = tempLeftovers[bestMatchIndex];
          const remainingAfterUse = usedPiece - usedFromLeftover;
          tempLeftovers.splice(bestMatchIndex, 1);
          if (remainingAfterUse > EPSILON) {
            tempLeftovers.push(remainingAfterUse);
            tempLeftovers.sort((a, b) => b - a);
          }
        } else {
          console.warn(
            `[WARNING] (in-memory): Could not find suitable leftover piece to consume ${usedFromLeftover.toFixed(4)} for ${partNumber} (${finish}).`,
          );
        }
        partExtra.length_pieces = tempLeftovers;
      }
      // Ensure length_pieces is a list
      if (!Array.isArray(partExtra.length_pieces)) {
        partExtra.length_pieces = [];
      }
      partExtra.quantity = 0.0;
    }

    const leftoverGenerated =
      materialImpactDetails.leftover_generated_qty_or_length ?? 0.0;
    const allNewLeftovers =
      materialImpactDetails.all_new_leftovers ?? [];

    // Get minPurchaseLength for validation
    let minPurchaseLength = 24.0;
    if (partNumber && partNumber !== 'N/A') {
      const partInfo = partsData[partNumber];
      if (partInfo) {
        const lengthStr = partInfo['Length'] ?? '';
        minPurchaseLength = parseLengthToFeet(lengthStr) || 24.0;
      }
    }

    // Ensure length_pieces is a list before appending
    if (!Array.isArray(partExtra.length_pieces)) {
      partExtra.length_pieces = [];
    }

    if (allNewLeftovers.length > 0) {
      for (const leftover of allNewLeftovers) {
        if (leftover > EPSILON && leftover < minPurchaseLength - EPSILON) {
          partExtra.length_pieces.push(leftover);
        }
      }
    } else if (leftoverGenerated > EPSILON) {
      if (leftoverGenerated < minPurchaseLength - EPSILON) {
        partExtra.length_pieces.push(leftoverGenerated);
      }
    }

    // Sort the list after adding new leftovers (ascending, matching Python .sort())
    if (Array.isArray(partExtra.length_pieces)) {
      partExtra.length_pieces.sort((a, b) => a - b);
    }
  } else if (typeProcessedAs === 'accessory') {
    const currentQty = partExtra.quantity ?? 0;
    const netChange =
      (materialImpactDetails.leftover_generated_qty_or_length ?? 0.0) -
      (materialImpactDetails.used_from_leftover_qty_or_length ?? 0.0);
    partExtra.quantity = Math.max(
      0,
      Math.round((currentQty + netChange) * 10000) / 10000,
    );
    partExtra.length_pieces = [];
  }

  materialsDict[extraMaterialsKey] = partExtra;
}

// ---------------------------------------------------------------------------
// Reverse material impact
// ---------------------------------------------------------------------------

/**
 * Reverses the material impact of a deleted elevation on the extra materials
 * inventory (in-memory version).
 */
export function reverseMaterialImpact(
  elevationMaterialImpacts: MaterialImpactDetails[] | null | undefined,
  extraMaterials: Record<string, ExtraMaterial>,
): void {
  if (!elevationMaterialImpacts) return;

  for (const impact of elevationMaterialImpacts) {
    const partNumber = impact.part_number;
    const typeProcessedAs = impact.type_processed_as;
    const finish = impact.finish;

    if (!partNumber || partNumber === 'N/A - Manual') continue;

    // Construct the key for extra materials based on part number and finish (for profiles)
    let extraMaterialsKey = partNumber;
    if (typeProcessedAs === 'profile' && finish) {
      extraMaterialsKey = `${partNumber}-${finish.toLowerCase()}`;
    }

    let partExtra: ExtraMaterial = extraMaterials[extraMaterialsKey]
      ? {
          quantity: extraMaterials[extraMaterialsKey].quantity ?? 0,
          length_pieces: [
            ...(extraMaterials[extraMaterialsKey].length_pieces ?? []),
          ],
        }
      : { quantity: 0, length_pieces: [] };

    const leftoverGenerated =
      impact.leftover_generated_qty_or_length ?? 0.0;
    const usedFromLeftover =
      impact.used_from_leftover_qty_or_length ?? 0.0;

    if (typeProcessedAs === 'profile') {
      const allNewLeftovers = impact.all_new_leftovers ?? [];
      const leftoverPiecesConsumed = impact.leftover_pieces_consumed ?? [];

      // When reversing, REMOVE all generated leftover pieces from inventory.
      if (allNewLeftovers.length > 0) {
        // Multi-piece leftovers from bin-packing — remove each one
        const tempLeftovers = [...(partExtra.length_pieces ?? [])];
        for (const piece of allNewLeftovers) {
          if (piece <= EPSILON) continue;
          let removed = false;
          for (let i = 0; i < tempLeftovers.length; i++) {
            if (Math.abs(tempLeftovers[i] - piece) < EPSILON) {
              tempLeftovers.splice(i, 1);
              removed = true;
              break;
            }
          }
          if (!removed) {
            console.warn(
              `[WARNING] Generated leftover '${piece.toFixed(4)} ft' for ${partNumber} (${finish}) not found in current inventory for reversal.`,
            );
          }
        }
        partExtra.length_pieces = tempLeftovers;
      } else if (leftoverGenerated > EPSILON) {
        // Single leftover fallback
        let removed = false;
        const tempLeftovers = [...(partExtra.length_pieces ?? [])];
        for (let i = 0; i < tempLeftovers.length; i++) {
          if (Math.abs(tempLeftovers[i] - leftoverGenerated) < EPSILON) {
            tempLeftovers.splice(i, 1);
            removed = true;
            break;
          }
        }
        if (!removed) {
          console.warn(
            `[WARNING] Generated leftover '${leftoverGenerated.toFixed(4)} ft' for ${partNumber} (${finish}) not found in current inventory for reversal.`,
          );
        }
        partExtra.length_pieces = tempLeftovers;
      }

      // When reversing, RESTORE all consumed leftover pieces back into inventory.
      if (leftoverPiecesConsumed.length > 0) {
        if (!Array.isArray(partExtra.length_pieces)) {
          partExtra.length_pieces = [];
        }
        for (const consumed of leftoverPiecesConsumed) {
          // Bay-width format: { original_length, used_length }
          if (typeof consumed === 'object' && !Array.isArray(consumed) && 'original_length' in consumed) {
            partExtra.length_pieces.push((consumed as { original_length: number }).original_length);
          }
          // Tuple format: [index, length_used] — restore the used length
          else if (Array.isArray(consumed) && consumed.length === 2) {
            partExtra.length_pieces.push((consumed as [number, number])[1]);
          }
        }
      } else if (usedFromLeftover > EPSILON) {
        // Single consumed fallback
        if (!Array.isArray(partExtra.length_pieces)) {
          partExtra.length_pieces = [];
        }
        partExtra.length_pieces.push(usedFromLeftover);
      }

      partExtra.length_pieces.sort((a, b) => a - b);
      partExtra.quantity = 0.0;
    } else if (typeProcessedAs === 'accessory') {
      const currentQty = partExtra.quantity ?? 0;
      const reverseNetChange = usedFromLeftover - leftoverGenerated;
      partExtra.quantity = Math.max(
        0,
        Math.round((currentQty + reverseNetChange) * 10000) / 10000,
      );
      partExtra.length_pieces = [];
    }

    extraMaterials[extraMaterialsKey] = partExtra;
  }

  // Clean up empty/ghost entries so stale finish keys don't persist
  // (e.g. when finish changes from Black→Clear, the old "-black" key is emptied)
  for (const key of Object.keys(extraMaterials)) {
    const entry = extraMaterials[key];
    const hasLength = entry.length_pieces && entry.length_pieces.length > 0;
    const hasQty = entry.quantity > 0;
    if (!hasLength && !hasQty) {
      delete extraMaterials[key];
    }
  }
}
