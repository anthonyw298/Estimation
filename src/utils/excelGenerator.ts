// Comprehensive Excel generator matching original Python functionality
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { ElevationData } from './storage';
import { getPriceByPart, getMultiplier, MaterialImpact, ExtraMaterials, loadExtraMaterials, saveExtraMaterials, getUnitPriceByPart, parseLengthToFeet, isBayWidthPart } from './pricing';
import { PART_NUMBER_MAP } from '../data/partNumber';
import { partsData } from '../data/partsData';

const COL_A = 1;
const COL_B = 2;
const COL_E = 5;
const PRICE_COL = 9;

interface SectionTotals {
  original: number;
  discounted: number;
}

function formatDoorSummary(calculatedOutputs: any[]): string {
  if (!calculatedOutputs) return "None";
  const doorLines: string[] = [];
  for (const item of calculatedOutputs) {
    if (item.type?.toLowerCase() === 'doors' && item.manual) {
      const quantity = item.quantity || 1;
      const style = item.Style || '';
      const price = item.price || 0.0;
      const hardware = item.hardware || {};
      if (!style || style.toLowerCase() === 'unknown') continue;
      const enabledHw = Object.entries(hardware)
        .filter(([_, enabled]) => enabled)
        .map(([hw]) => hw);
      if (enabledHw.length > 0) {
        doorLines.push(`${quantity} x ${style} Door ($${price.toFixed(2)})\n  with: ${enabledHw.join(', ')}`);
      } else {
        doorLines.push(`${quantity} x ${style} Door ($${price.toFixed(2)})`);
      }
    }
  }
  return doorLines.length > 0 ? doorLines.join('; \n') : 'None';
}

async function writeOutputSection(
  worksheet: ExcelJS.Worksheet,
  title: string,
  items: any[],
  colE: number,
  elevationFinish: string,
  systemTotalRef: { value: number },
  originalSystemTotalRef: { value: number },
  startOutputRow: number,
  currentExtraMaterials: ExtraMaterials,
  projectName: string,
  multiplier: number
): Promise<[number, MaterialImpact[], SectionTotals]> {
  if (!items || items.length === 0) {
    return [startOutputRow, [], { original: 0.0, discounted: 0.0 }];
  }

  let currentRow = startOutputRow;
  const titleCell = worksheet.getCell(currentRow, colE);
  titleCell.value = title;
  titleCell.font = { bold: true, size: 12 };

  const headers = ["Description", "Part Number", "Total Quantity Required", "Total List Cost", "Discounted Total List Cost"];
  for (let i = 0; i < headers.length; i++) {
    const headerCell = worksheet.getCell(currentRow + 1, colE + i);
    headerCell.value = headers[i];
    headerCell.font = { bold: true };
    headerCell.border = {
      bottom: { style: 'thin' }
    };
  }
  currentRow += 2;

  const sectionMaterialImpacts: MaterialImpact[] = [];
  let sectionOriginalTotal = 0.0;
  let sectionDiscountedTotal = 0.0;

  for (const item of items) {
    const qtyRaw = item.quantity || 0;
    const pn = item.part_number || '';
    const manual = item.manual || false;
    const desc = (item.description || '').trim();
    const isProfile = pn in (PART_NUMBER_MAP.profiles || {});
    const isGasket = desc.toLowerCase().includes('gasket') || ['E2-0052', 'E2-0053', 'E2-0065'].includes(pn);
    const isAccessory = pn in (PART_NUMBER_MAP.accessories || {}) || item.type?.toLowerCase() === 'accessory';
    const isGlass = pn === 'GLASS_AREA' || item.type?.toLowerCase() === 'glass';

    const individualQuantities = Array.isArray(qtyRaw) ? qtyRaw : [qtyRaw];
    const qtySum = individualQuantities.reduce((sum, q) => sum + (typeof q === 'number' ? q : parseFloat(q.toString()) || 0), 0);

    const unitType = isProfile || isGasket ? 'ft' : isAccessory ? 'pcs' : item.unit || (isGlass ? 'sqft' : 'pcs');
    const displayUnit = unitType;

    let displayQtyString: string;
    if (Array.isArray(qtyRaw)) {
      if (qtyRaw.length > 1 && qtyRaw.every(x => x === qtyRaw[0])) {
        displayQtyString = `${qtyRaw[0].toFixed(2)} ${displayUnit} x ${qtyRaw.length}`;
      } else {
        displayQtyString = qtyRaw.map(q => `${q.toFixed(2)} ${displayUnit}`).join(', ');
      }
    } else {
      displayQtyString = `${qtyRaw.toFixed(2)} ${displayUnit}`;
    }

    let itemTotalCostForDisplay = 0.0;
    let originalItemTotalCost = 0.0;

    // Check if this is a bay width part that should process the list as a whole
    const isBayWidthItem = isBayWidthPart(pn, qtyRaw, desc);
    
    // For bay width parts with a list, process the entire list at once for optimization
    // This matches the Python version's logic (lines 258-273)
    if (isBayWidthItem && Array.isArray(qtyRaw) && qtyRaw.length > 1) {
      const useGroup = isGasket;
      const [totalPrice, calculatedUnitType, materialImpact] = getPriceByPart(
        pn,
        qtyRaw,
        elevationFinish,
        currentExtraMaterials,
        false,
        useGroup,
        projectName,
        desc
      );
      
      itemTotalCostForDisplay = totalPrice !== null ? totalPrice : 0.0;
      originalItemTotalCost = totalPrice !== null ? totalPrice : 0.0;
      
      if (materialImpact) {
        sectionMaterialImpacts.push(materialImpact);
        // Apply material impact in memory for profiles
        if (materialImpact.type_processed_as === 'profile') {
          const key = elevationFinish ? `${pn}-${elevationFinish.toLowerCase()}` : pn;
          if (!currentExtraMaterials[key]) {
            currentExtraMaterials[key] = { quantity: 0, length_pieces: [] };
          }
          // Handle leftover pieces consumption for bay width parts
          const leftover = materialImpact.leftover_generated_qty_or_length || 0;
          if (leftover > 0.0001) {
            currentExtraMaterials[key].length_pieces.push(leftover);
            currentExtraMaterials[key].length_pieces.sort((a, b) => b - a);
          }
          // Remove used leftover pieces
          const usedFromLeftover = materialImpact.used_from_leftover_qty_or_length || 0;
          if (usedFromLeftover > 0.0001) {
            const leftovers = currentExtraMaterials[key].length_pieces;
            for (let i = leftovers.length - 1; i >= 0; i--) {
              if (leftovers[i] >= usedFromLeftover - 0.0001) {
                const remaining = leftovers[i] - usedFromLeftover;
                leftovers.splice(i, 1);
                if (remaining > 0.0001) {
                  leftovers.push(remaining);
                  leftovers.sort((a, b) => b - a);
                }
                break;
              }
            }
            currentExtraMaterials[key].length_pieces = leftovers;
          }
        }
      }
    } else {
      // Standard processing: iterate through each quantity
      for (const singleQty of individualQuantities) {
        let totalItemPriceSingleCut = 0.0;
        let calculatedUnitType = unitType;

        if (manual) {
          if (pn && pn !== 'N/A') {
            const [priceCalculated, unitCalculated, materialImpact] = getPriceByPart(
              pn,
              singleQty,
              elevationFinish,
              currentExtraMaterials,
              false,
              true,
              projectName,
              desc
            );
            totalItemPriceSingleCut = priceCalculated !== null ? priceCalculated : (item.price || 0.0) * singleQty;
            calculatedUnitType = isProfile || isAccessory ? unitType : (unitCalculated || item.unit || 'pcs');
            if (materialImpact) {
              sectionMaterialImpacts.push(materialImpact);
              // Apply material impact in memory
              if (materialImpact.type_processed_as === 'profile') {
                const key = elevationFinish ? `${pn}-${elevationFinish.toLowerCase()}` : pn;
                if (!currentExtraMaterials[key]) {
                  currentExtraMaterials[key] = { quantity: 0, length_pieces: [] };
                }
                const leftover = materialImpact.leftover_generated_qty_or_length || 0;
                if (leftover > 0.0001) {
                  currentExtraMaterials[key].length_pieces.push(leftover);
                  currentExtraMaterials[key].length_pieces.sort((a, b) => b - a);
                }
                // Remove used leftover pieces
                const usedFromLeftover = materialImpact.used_from_leftover_qty_or_length || 0;
                if (usedFromLeftover > 0.0001) {
                  const leftovers = currentExtraMaterials[key].length_pieces;
                  for (let i = leftovers.length - 1; i >= 0; i--) {
                    if (leftovers[i] >= usedFromLeftover - 0.0001) {
                      const remaining = leftovers[i] - usedFromLeftover;
                      leftovers.splice(i, 1);
                      if (remaining > 0.0001) {
                        leftovers.push(remaining);
                        leftovers.sort((a, b) => b - a);
                      }
                      break;
                    }
                  }
                  currentExtraMaterials[key].length_pieces = leftovers;
                }
              }
            }
          } else {
            totalItemPriceSingleCut = (item.price || 0.0) * singleQty;
            calculatedUnitType = item.unit || 'pcs';
          }
        } else {
          const useGroup = isGasket;
          const [totalPrice, unitFromPricing, materialImpact] = getPriceByPart(
            pn,
            singleQty,
            elevationFinish,
            currentExtraMaterials,
            false,
            useGroup,
            projectName
          );
          totalItemPriceSingleCut = totalPrice !== null ? totalPrice : 0.0;
          calculatedUnitType = isProfile || isGasket || isAccessory ? unitType : (unitFromPricing || item.unit || 'pcs');
          if (materialImpact) {
            sectionMaterialImpacts.push(materialImpact);
            // Apply material impact in memory
            if (materialImpact.type_processed_as === 'profile') {
              const key = elevationFinish ? `${pn}-${elevationFinish.toLowerCase()}` : pn;
              if (!currentExtraMaterials[key]) {
                currentExtraMaterials[key] = { quantity: 0, length_pieces: [] };
              }
              const leftover = materialImpact.leftover_generated_qty_or_length || 0;
              if (leftover > 0.0001) {
                currentExtraMaterials[key].length_pieces.push(leftover);
                currentExtraMaterials[key].length_pieces.sort((a, b) => b - a);
              }
              // Remove used leftover pieces
              const usedFromLeftover = materialImpact.used_from_leftover_qty_or_length || 0;
              if (usedFromLeftover > 0.0001) {
                const leftovers = currentExtraMaterials[key].length_pieces;
                for (let i = leftovers.length - 1; i >= 0; i--) {
                  if (leftovers[i] >= usedFromLeftover - 0.0001) {
                    const remaining = leftovers[i] - usedFromLeftover;
                    leftovers.splice(i, 1);
                    if (remaining > 0.0001) {
                      leftovers.push(remaining);
                      leftovers.sort((a, b) => b - a);
                    }
                    break;
                  }
                }
                currentExtraMaterials[key].length_pieces = leftovers;
              }
            }
          }
        }

        itemTotalCostForDisplay += totalItemPriceSingleCut;
        originalItemTotalCost += totalItemPriceSingleCut;
      }
    }

    if (isProfile || isGasket || isAccessory) {
      itemTotalCostForDisplay *= multiplier;
      if (qtySum > 0) {
        item.price = itemTotalCostForDisplay / qtySum;
      }
    }

    systemTotalRef.value += itemTotalCostForDisplay;
    originalSystemTotalRef.value += originalItemTotalCost;
    sectionOriginalTotal += originalItemTotalCost;
    sectionDiscountedTotal += itemTotalCostForDisplay;

    worksheet.getCell(currentRow, colE).value = item.description || '';
    worksheet.getCell(currentRow, colE + 1).value = pn || 'N/A';
    worksheet.getCell(currentRow, colE + 2).value = displayQtyString;
    worksheet.getCell(currentRow, colE + 3).value = originalItemTotalCost;
    worksheet.getCell(currentRow, colE + 3).numFmt = '$#,##0.00';
    worksheet.getCell(currentRow, colE + 4).value = itemTotalCostForDisplay;
    worksheet.getCell(currentRow, colE + 4).numFmt = '$#,##0.00';
    currentRow += 1;
  }

  // Add Section Totals
  const titleMapping: Record<string, string> = {
    "PROFILES": "Profile",
    "ACCESSORIES": "Accessory",
    "GASKETS": "Gasket",
    "DOORS": "Door",
    "GLASS": "Glass",
    "LABOR": "Labor"
  };
  const titleLabel = titleMapping[title.toUpperCase()] || title;
  const totalLabel = `Total ${titleLabel} Cost`;
  
  worksheet.getCell(currentRow, colE + 2).value = totalLabel;
  worksheet.getCell(currentRow, colE + 2).font = { bold: true };
  worksheet.getCell(currentRow, colE + 3).value = sectionOriginalTotal;
  worksheet.getCell(currentRow, colE + 3).numFmt = '$#,##0.00';
  worksheet.getCell(currentRow, colE + 3).font = { bold: true };
  worksheet.getCell(currentRow, colE + 4).value = sectionDiscountedTotal;
  worksheet.getCell(currentRow, colE + 4).numFmt = '$#,##0.00';
  worksheet.getCell(currentRow, colE + 4).font = { bold: true };

  // Add top border for totals row
  for (let col = colE; col < colE + 5; col++) {
    worksheet.getCell(currentRow, col).border = {
      top: { style: 'thin' }
    };
  }

  return [currentRow + 1, sectionMaterialImpacts, { original: sectionOriginalTotal, discounted: sectionDiscountedTotal }];
}

export async function generateExcelReport(
  projectName: string,
  elevations: Record<string, ElevationData>
): Promise<void> {
  if (!projectName) {
    throw new Error('Project name is required');
  }
  if (!elevations || Object.keys(elevations).length === 0) {
    throw new Error('No elevations found to generate report');
  }

  console.log('Generating Excel report for project:', projectName);
  console.log('Elevations:', Object.keys(elevations));
  
  const workbook = new ExcelJS.Workbook();
  
  // Calculate multiplier based on grand total
  // This should match the Python version's logic in create_summary_sheet
  let fullRunningGrandTotal = 0.0;
  for (const elev of Object.values(elevations)) {
    const elevationFinish = (elev.finish || '').toLowerCase();
    for (const output of elev.calculated_outputs || []) {
      let qty: number | number[] = output.quantity || 0;
      // Sum up list quantities for multiplier calculation (matching Python line 386)
      if (Array.isArray(qty)) {
        qty = qty.reduce((sum: number, q: any) => sum + (typeof q === 'number' ? q : parseFloat(q.toString()) || 0), 0);
      }
      const manual = output.manual || false;
      const part = (output.part_number || '').trim();
      const itemType = (output.type || '').toLowerCase();
      
      let price = 0.0;
      if (manual || part === 'GLASS_AREA' || ['glass', 'joints_fab_labor', 'door', 'doors'].includes(itemType)) {
        price = (output.price || 0.0) * (typeof qty === 'number' ? qty : qty.reduce((s, q) => s + q, 0));
      } else if (part && part !== 'N/A') {
        const [calculatedPrice] = getPriceByPart(
          part,
          qty,
          elevationFinish,
          undefined,
          true, // summary = true
          false, // group = false (unless manual, but manual items are handled above)
          projectName
        );
        price = calculatedPrice !== null ? calculatedPrice : 0.0;
      }
      fullRunningGrandTotal += price;
    }
  }
  const multiplier = getMultiplier(fullRunningGrandTotal);

  // Load extra materials once for the entire project (shared across all elevations)
  // This will be updated as we process each elevation
  const currentExtraMaterials = loadExtraMaterials(projectName);

  const sortedElevNames = Object.keys(elevations).sort();
  
  if (sortedElevNames.length === 0) {
    throw new Error('No elevations found to generate report');
  }

  for (const elevName of sortedElevNames) {
    const worksheet = workbook.addWorksheet(elevName);
    const elevData = elevations[elevName];
    
    if (!elevData) {
      console.warn(`Skipping elevation ${elevName}: data is missing`);
      continue;
    }

    // Format custom bay dimensions
    const customBayWidthsStr = elevData.custom_bay_widths && elevData.custom_bay_widths.length > 0
      ? elevData.custom_bay_widths.map(w => `${w.toFixed(2)} in`).join(', ')
      : 'Equal distribution';
    const customBayHeightsStr = elevData.custom_bay_heights && elevData.custom_bay_heights.length > 0
      ? elevData.custom_bay_heights.map(h => `${h.toFixed(2)} in`).join(', ')
      : 'Equal distribution';

    // Input data section
    const inputData = [
      ["System Input", elevData.system || ''],
      ["Finish", elevData.finish || ''],
      ["Elevation Type", elevName],
      ["Total Count", elevData.total_count || 0],
      ["Bays Wide", elevData.bays_wide || ''],
      ["Bays Tall", elevData.bays_tall || ''],
      ["Custom Bay Widths", customBayWidthsStr],
      ["Custom Bay Heights", customBayHeightsStr],
      ["Opening Width", `${elevData.opening_width_inches.toFixed(2)} in`],
      ["Opening Height", `${elevData.opening_height_inches.toFixed(2)} in`],
      ["Sq Ft per Type", `${elevData.sqft_per_type.toFixed(2)} sqft`],
      ["Total Sq Ft", `${elevData.total_sqft.toFixed(2)} sqft`],
      ["Perimeter Ft", `${elevData.perimeter_ft.toFixed(2)} ft`],
      ["Total Perimeter Ft", `${elevData.total_perimeter_ft.toFixed(2)} ft`],
      ["Doors", formatDoorSummary(elevData.calculated_outputs || [])]
    ];

    let currentExcelRow = 1;
    const thinBorder = {
      left: { style: 'thin' as const },
      right: { style: 'thin' as const },
      top: { style: 'thin' as const },
      bottom: { style: 'thin' as const }
    };

    for (let i = 0; i < inputData.length; i++) {
      const [header, value] = inputData[i];
      const headerCell = worksheet.getCell(currentExcelRow + i, COL_A);
      headerCell.value = header;
      headerCell.font = { bold: true };
      headerCell.border = thinBorder;

      const valueCell = worksheet.getCell(currentExcelRow + i, COL_B);
      valueCell.value = value as string | number;
      valueCell.border = thinBorder;
      if (['Total Count', 'Bays Wide', 'Bays Tall'].includes(header)) {
        valueCell.alignment = { horizontal: 'left' };
      }
    }

    // Categorize outputs
    const profilesForSection: any[] = [];
    const accessoriesForSection: any[] = [];
    const gasketsForSection: any[] = [];
    const otherItemsForSection: any[] = [];

    const currentElevationFinish = elevData.finish || '';

    for (const item of elevData.calculated_outputs || []) {
      const pn = item.part_number || '';
      const manual = item.manual || false;
      const desc = (item.description || '').trim();
      const isGasket = desc.toLowerCase().includes('gasket') || ['E2-0052', 'E2-0053', 'E2-0065'].includes(pn);

      if (pn && pn !== 'N/A') {
        if (manual) {
          otherItemsForSection.push(item);
        } else if (pn in (PART_NUMBER_MAP.profiles || {})) {
          profilesForSection.push(item);
        } else if (isGasket) {
          gasketsForSection.push(item);
        } else if (pn in (PART_NUMBER_MAP.accessories || {}) || item.type?.toLowerCase() === 'accessory') {
          accessoriesForSection.push(item);
        } else {
          otherItemsForSection.push(item);
        }
      } else {
        otherItemsForSection.push(item);
      }
    }

    const systemTotalForThisBlock = { value: 0.0 };
    const originalSystemTotalForThisBlock = { value: 0.0 };
    // Use the shared extraMaterials object (already loaded above)

    let outputSectionCurrentRow = 1;

    // Write PROFILES section
    const [nextRowAfterProfiles, , profileTotals] = await writeOutputSection(
      worksheet, "PROFILES", profilesForSection, COL_E, currentElevationFinish,
      systemTotalForThisBlock, originalSystemTotalForThisBlock, outputSectionCurrentRow,
      currentExtraMaterials, projectName, multiplier
    );
    const profileOriginalTotal = profileTotals.original;
    const profileDiscountedTotal = profileTotals.discounted;

    // Write ACCESSORIES section
    const [nextRowAfterAccessories, , accessoryTotals] = await writeOutputSection(
      worksheet, "ACCESSORIES", accessoriesForSection, COL_E, currentElevationFinish,
      systemTotalForThisBlock, originalSystemTotalForThisBlock, nextRowAfterProfiles,
      currentExtraMaterials, projectName, multiplier
    );
    const accessoryOriginalTotal = accessoryTotals.original;
    const accessoryDiscountedTotal = accessoryTotals.discounted;

    // Write GASKETS section
    const [nextRowAfterGaskets, , gasketTotals] = await writeOutputSection(
      worksheet, "GASKETS", gasketsForSection, COL_E, currentElevationFinish,
      systemTotalForThisBlock, originalSystemTotalForThisBlock, nextRowAfterAccessories,
      currentExtraMaterials, projectName, multiplier
    );
    const gasketOriginalTotal = gasketTotals.original;
    const gasketDiscountedTotal = gasketTotals.discounted;

    // Group other items
    const groupedOtherMisc: Record<string, any[]> = {};
    for (const item of otherItemsForSection) {
      const itemType = (item.type || 'MISCELLANEOUS ITEMS').toUpperCase();
      if (!groupedOtherMisc[itemType]) {
        groupedOtherMisc[itemType] = [];
      }
      groupedOtherMisc[itemType].push(item);
    }

    let currentSectionRow = nextRowAfterGaskets;
    let glassOriginalTotal = 0.0;
    let glassDiscountedTotal = 0.0;
    let fabricationOriginalTotal = 0.0;
    let fabricationDiscountedTotal = 0.0;

    for (const [grpTitle, grpItems] of Object.entries(groupedOtherMisc)) {
      const [nextRowAfterGroup, , groupTotals] = await writeOutputSection(
        worksheet, grpTitle, grpItems, COL_E, currentElevationFinish,
        systemTotalForThisBlock, originalSystemTotalForThisBlock, currentSectionRow,
        currentExtraMaterials, projectName, multiplier
      );

      if (grpTitle === 'GLASS' || grpItems.some(item => item.part_number === 'GLASS_AREA' || item.type?.toLowerCase() === 'glass')) {
        glassOriginalTotal += groupTotals.original;
        glassDiscountedTotal += groupTotals.discounted;
      } else {
        fabricationOriginalTotal += groupTotals.original;
        fabricationDiscountedTotal += groupTotals.discounted;
      }

      currentSectionRow = nextRowAfterGroup;
    }

    // Cost breakdown summary
    const spacingRows = 1;
    for (let blankRow = 1; blankRow <= spacingRows; blankRow++) {
      worksheet.getCell(currentSectionRow + blankRow, COL_A).value = '';
    }

    const costSummaryRow = currentSectionRow + spacingRows + 1;
    const headerCol = PRICE_COL - 2;
    const costPerElevCol = PRICE_COL - 1;
    const totalElevCostCol = PRICE_COL;

    // Headers
    worksheet.getCell(costSummaryRow, headerCol).value = 'COST/ELEVATION';
    worksheet.getCell(costSummaryRow, headerCol).font = { bold: true };
    worksheet.getCell(costSummaryRow, costPerElevCol).value = 'COST/ELEVATION';
    worksheet.getCell(costSummaryRow, costPerElevCol).font = { bold: true };
    worksheet.getCell(costSummaryRow, totalElevCostCol).value = 'TOTAL ELEVATION COST';
    worksheet.getCell(costSummaryRow, totalElevCostCol).font = { bold: true };

    for (const col of [headerCol, costPerElevCol, totalElevCostCol]) {
      worksheet.getCell(costSummaryRow, col).border = {
        bottom: { style: 'thin' }
      };
    }

    let costSummaryCurrentRow = costSummaryRow + 1;
    const totalCount = elevData.total_count || 1;

    // Profile Costs
    worksheet.getCell(costSummaryCurrentRow, headerCol).value = 'PROFILE COSTS';
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).value = profileDiscountedTotal;
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).numFmt = '$#,##0.00';
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).value = profileDiscountedTotal * totalCount;
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).numFmt = '$#,##0.00';
    costSummaryCurrentRow += 1;

    // Accessory Costs
    worksheet.getCell(costSummaryCurrentRow, headerCol).value = 'ACCESSORY COSTS';
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).value = accessoryDiscountedTotal;
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).numFmt = '$#,##0.00';
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).value = accessoryDiscountedTotal * totalCount;
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).numFmt = '$#,##0.00';
    costSummaryCurrentRow += 1;

    // Gasket Costs
    worksheet.getCell(costSummaryCurrentRow, headerCol).value = 'GASKET COSTS';
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).value = gasketDiscountedTotal;
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).numFmt = '$#,##0.00';
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).value = gasketDiscountedTotal * totalCount;
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).numFmt = '$#,##0.00';
    costSummaryCurrentRow += 1;

    // Glass Costs
    worksheet.getCell(costSummaryCurrentRow, headerCol).value = 'GLASS COSTS';
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).value = glassDiscountedTotal;
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).numFmt = '$#,##0.00';
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).value = glassDiscountedTotal * totalCount;
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).numFmt = '$#,##0.00';
    costSummaryCurrentRow += 1;

    // Fabrication Costs
    worksheet.getCell(costSummaryCurrentRow, headerCol).value = 'FABRICATION COSTS';
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).value = fabricationDiscountedTotal;
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).numFmt = '$#,##0.00';
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).value = fabricationDiscountedTotal * totalCount;
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).numFmt = '$#,##0.00';
    costSummaryCurrentRow += 1;

    // Separator line
    for (const col of [headerCol, costPerElevCol, totalElevCostCol]) {
      worksheet.getCell(costSummaryCurrentRow, col).border = {
        top: { style: 'thin' }
      };
    }
    costSummaryCurrentRow += 1;

    // Total Costs
    const totalCostPerElev = profileDiscountedTotal + accessoryDiscountedTotal + gasketDiscountedTotal + glassDiscountedTotal + fabricationDiscountedTotal;
    const totalElevationCost = totalCostPerElev * totalCount;

    worksheet.getCell(costSummaryCurrentRow, headerCol).value = `${elevName} TOTAL COSTS`;
    worksheet.getCell(costSummaryCurrentRow, headerCol).font = { bold: true };
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).value = totalCostPerElev;
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).numFmt = '$#,##0.00';
    worksheet.getCell(costSummaryCurrentRow, costPerElevCol).font = { bold: true };
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).value = totalElevationCost;
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).numFmt = '$#,##0.00';
    worksheet.getCell(costSummaryCurrentRow, totalElevCostCol).font = { bold: true };
    costSummaryCurrentRow += 1;

    // Note
    worksheet.getCell(costSummaryCurrentRow, headerCol).value = '*Note - Elevation costs based on discounted material costs';
    worksheet.getCell(costSummaryCurrentRow, headerCol).font = { italic: true, size: 10 };

    // Auto-fit columns
    worksheet.columns.forEach((column, index) => {
      if (column && column.eachCell) {
        let maxLength = 0;
        column.eachCell({ includeEmpty: false }, (cell) => {
          const cellValue = cell.value?.toString() || '';
          maxLength = Math.max(maxLength, cellValue.length);
        });
        if (maxLength > 0) {
          worksheet.getColumn(index + 1).width = Math.min(maxLength + 2, 50);
        }
      }
    });
  }

  // Save extra materials state (save the updated state from all elevations)
  saveExtraMaterials(projectName, currentExtraMaterials);

  // Create comprehensive Summary sheet matching original
  await createSummarySheet(workbook, elevations, projectName, currentExtraMaterials, multiplier);

  // Generate file - match original naming convention
  // Original saves to: reports/{project_name}_{timestamp}.xlsx
  console.log('Writing workbook to buffer...');
  let buffer: ArrayBuffer;
  try {
    buffer = await workbook.xlsx.writeBuffer();
    console.log('Buffer created, size:', buffer.byteLength);
  } catch (writeErr: any) {
    console.error('Error writing workbook buffer:', writeErr);
    throw new Error(`Failed to generate Excel file: ${writeErr.message || writeErr.toString()}`);
  }
  
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  
  // Create timestamp in format: YYYYMMDD_HHMMSS (matching Python datetime.strftime)
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  const timestamp = `${year}${month}${day}_${hours}${minutes}${seconds}`;
  
  // Match original naming: reports/{project_name}_{timestamp}.xlsx
  const fileName = `${projectName}_${timestamp}.xlsx`;
  
  console.log('Saving file:', fileName);
  console.log('Blob size:', blob.size);
  
  // Always use file-saver for now (more reliable across browsers)
  // File System Access API requires user interaction and may be blocked
  try {
    saveAs(blob, fileName);
    console.log('File download initiated successfully');
  } catch (downloadErr: any) {
    console.error('Error downloading file:', downloadErr);
    throw new Error(`Failed to download file: ${downloadErr.message || downloadErr.toString()}`);
  }
  
  // Optional: Try File System Access API if user wants to choose location
  // This is commented out because it requires user interaction and may be blocked
  /*
  try {
    if ('showSaveFilePicker' in window) {
      try {
        const fileHandle = await (window as any).showSaveFilePicker({
          suggestedName: fileName,
          types: [{
            description: 'Excel files',
            accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
          }]
        });
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        console.log('File saved using File System Access API');
        return;
      } catch (pickerErr: any) {
        if (pickerErr.name === 'AbortError') {
          console.log('User cancelled file save');
          return; // Don't throw, just return silently
        }
        console.log('File System Access API error, using download instead:', pickerErr);
      }
    }
  } catch (err: any) {
    console.log('File System Access API not available, using download:', err);
  }
  */
}

async function createSummarySheet(
  workbook: ExcelJS.Workbook,
  elevations: Record<string, ElevationData>,
  projectName: string,
  extraMaterials: ExtraMaterials,
  multiplier: number
): Promise<void> {
  const summarySheet = workbook.addWorksheet('Summary');

  // Step 1: Aggregate quantities and prices across all elevations, grouped by category
  const categories: Record<string, any[]> = {
    'PROFILES': [],
    'ACCESSORIES': [],
    'GASKETS': [],
    'DOORS': [],
    'GLASS': [],
    'LABOR': []
  };

  for (const [elevKey, elev] of Object.entries(elevations)) {
    const elevationFinish = (elev.finish || '').toLowerCase();
    for (const output of elev.calculated_outputs || []) {
      const part = (output.part_number || '').trim();
      const desc = (output.description || '').trim();
      const manual = output.manual || false;
      const qty = output.quantity || 0;
      const qtyForAggregation = Array.isArray(qty) ? qty.reduce((sum, q) => sum + (typeof q === 'number' ? q : parseFloat(q.toString()) || 0), 0) : qty;
      
      const isProfile = part in (PART_NUMBER_MAP.profiles || {});
      const isGasket = desc.toLowerCase().includes('gasket') || ['E2-0052', 'E2-0053', 'E2-0065'].includes(part);
      const isAccessory = part in (PART_NUMBER_MAP.accessories || {}) || output.type?.toLowerCase() === 'accessory';
      const isGlass = part === 'GLASS_AREA' || output.type?.toLowerCase() === 'glass';
      const isJointsFabLabor = part === 'JOINTS_FAB_LABOR' || output.type?.toLowerCase() === 'joints_fab_labor' || desc.toLowerCase().includes('joints fabrication') || desc.toLowerCase().includes('fabrication labor');
      const isDoor = output.type?.toLowerCase() === 'door' || output.type?.toLowerCase() === 'doors';

      let category: string | null = null;
      if (isProfile) category = 'PROFILES';
      else if (isGasket) category = 'GASKETS';
      else if (isAccessory) category = 'ACCESSORIES';
      else if (isDoor) category = 'DOORS';
      else if (isGlass) category = 'GLASS';
      else if (isJointsFabLabor) category = 'LABOR';
      else continue;

      let key: string;
      let display: string;
      if (manual || isGlass || isJointsFabLabor || isDoor) {
        if (part && part !== 'N/A') {
          key = (isProfile || isGasket || isJointsFabLabor || isDoor || isGlass) && elevationFinish
            ? `MANUAL_${part}-${elevationFinish}`
            : `MANUAL_${part}`;
          display = (isProfile || isGasket || isJointsFabLabor || isDoor || isGlass) && elevationFinish
            ? `${desc} (${part} - ${elevationFinish})`
            : `${desc} (${part})`;
        } else {
          key = `MANUAL_NO_PN_${desc}`;
          display = desc;
        }
      } else {
        if ((isProfile || isGasket || isJointsFabLabor || isDoor || isGlass) && elevationFinish) {
          key = `${part}-${elevationFinish}`;
          display = `${part} (${elevationFinish})`;
        } else {
          key = part;
          display = part;
        }
      }

      categories[category].push({
        key,
        quantity: qtyForAggregation,
        quantity_list: Array.isArray(qty) ? qty : [qty],
        description: desc,
        display,
        part_number: part,
        manual,
        unit: isProfile || isGasket ? 'ft' : isAccessory ? 'pcs' : output.unit || (isGlass ? 'sqft' : 'pcs'),
        finish: (isProfile || isGasket || isJointsFabLabor || isDoor || isGlass) ? elevationFinish : '',
        is_glass: isGlass,
        is_joints_fab_labor: isJointsFabLabor,
        is_door: isDoor,
        is_gasket: isGasket,
        price: (manual || isGlass || isJointsFabLabor || isDoor) ? (output.price || 0.0) : 0.0
      });
    }
  }

  // Step 2: Aggregate items within each category by key
  for (const category of Object.keys(categories)) {
    const aggregatedMap: Record<string, any> = {};
    for (const item of categories[category]) {
      const k = item.key;
      if (k in aggregatedMap) {
        const existing = aggregatedMap[k];
        if (!existing.description && item.description) {
          existing.description = item.description;
        }
        if (existing.manual || existing.is_glass || existing.is_joints_fab_labor || existing.is_door) {
          const costExisting = (existing.price || 0.0) * existing.quantity;
          const costNew = (item.price || 0.0) * item.quantity;
          const totalQty = existing.quantity + item.quantity;
          existing.quantity = totalQty;
          existing.quantity_list = [...(existing.quantity_list || [existing.quantity]), ...(item.quantity_list || [item.quantity])];
          if (totalQty > 0) {
            existing.price = (costExisting + costNew) / totalQty;
          }
        } else {
          existing.quantity = existing.quantity + item.quantity;
          existing.quantity_list = [...(existing.quantity_list || [existing.quantity]), ...(item.quantity_list || [item.quantity])];
        }
      } else {
        item.quantity = typeof item.quantity === 'number' ? item.quantity : parseFloat(item.quantity.toString()) || 0;
        if (!item.quantity_list) item.quantity_list = [item.quantity];
        if (!item.description) item.description = item.display;
        aggregatedMap[k] = item;
      }
    }
    categories[category] = Object.values(aggregatedMap);
  }

  // Step 3: Calculate prices and prepare final data
  const finalSummaryData: any[] = [];
  let totalReusableCost = 0.0;
  let grandOriginalTotal = 0.0;
  let grandDiscountedTotal = 0.0;
  let grandResidualTotal = 0.0;

  for (const [category, items] of Object.entries(categories)) {
    for (const item of items) {
      const quantityAggregated = item.quantity;
      const manual = item.manual;
      const part = item.part_number;
      const display = item.display;
      const description = item.description || display;
      const isProfile = part in (PART_NUMBER_MAP.profiles || {});
      const isAccessory = part in (PART_NUMBER_MAP.accessories || {}) || item.type?.toLowerCase() === 'accessory';
      const isGasket = item.is_gasket || description.toLowerCase().includes('gasket') || ['E2-0052', 'E2-0053', 'E2-0065'].includes(part);
      const isGlass = item.is_glass;
      const isJointsFabLabor = item.is_joints_fab_labor;
      const isDoor = item.is_door;
      const itemFinish = item.finish;

      const displayUnit = isProfile || isGasket ? 'ft' : isAccessory ? 'pcs' : item.unit;
      let originalTotalCostForItem = 0.0;
      let totalCostForItem = 0.0;
      let calculatedUnitType = displayUnit;
      let reusableQtySum = 0.0;
      let reusablePct = 0.0;
      let reusableCost = 0.0;
      let reusableQtyDisplayString = 'N/A';

      if (manual || isGlass || isJointsFabLabor || isDoor) {
        const price = item.price || 0.0;
        const qtyFloat = typeof quantityAggregated === 'number' ? quantityAggregated : parseFloat(quantityAggregated.toString()) || 0;
        originalTotalCostForItem = price * qtyFloat;
        calculatedUnitType = item.unit || (isGlass ? 'sqft' : 'pcs');
      } else {
        const useGroup = isGasket;
        const [totalPrice, unitTypeFromPricing] = getPriceByPart(
          part,
          quantityAggregated,
          itemFinish,
          extraMaterials,
          true,
          useGroup,
          projectName
        );
        originalTotalCostForItem = totalPrice !== null ? totalPrice : 0.0;
        calculatedUnitType = isProfile ? 'ft' : isAccessory ? 'pcs' : (unitTypeFromPricing || item.unit || 'pcs');
      }

      if (isProfile || isGasket || isAccessory) {
        totalCostForItem = originalTotalCostForItem * multiplier;
      } else {
        totalCostForItem = originalTotalCostForItem;
      }

      if (part && part !== 'N/A' && (isProfile || isGasket || isAccessory)) {
        let extraMaterialsKeyForReuse = part;
        if ((isProfile || isGasket) && itemFinish) {
          extraMaterialsKeyForReuse = `${part}-${itemFinish}`;
        }

        const partData = extraMaterials[extraMaterialsKeyForReuse] || { quantity: 0, length_pieces: [] };
        if (partData.length_pieces && partData.length_pieces.length > 0) {
          const lengths = partData.length_pieces.filter((l: any) => typeof l === 'number' || !isNaN(parseFloat(l.toString())))
            .map((l: any) => typeof l === 'number' ? l : parseFloat(l.toString()));
          reusableQtySum = lengths.reduce((sum, l) => sum + l, 0);
          if (lengths.length > 0) {
            // Count occurrences
            const counter: Record<string, number> = {};
            lengths.forEach(l => {
              const key = l.toFixed(2);
              counter[key] = (counter[key] || 0) + 1;
            });
            const reuseLengthsFormatted = Object.entries(counter)
              .sort(([a], [b]) => parseFloat(b) - parseFloat(a))
              .map(([length, count]) => count > 1 ? `${length} ${displayUnit} x${count}` : `${length} ${displayUnit}`);
            reusableQtyDisplayString = reuseLengthsFormatted.join(', ');
          }
        } else {
          reusableQtySum = partData.quantity || 0.0;
          reusableQtyDisplayString = `${reusableQtySum.toFixed(2)} ${displayUnit}`;
        }

        const quantityAggregatedF = typeof quantityAggregated === 'number' ? quantityAggregated : parseFloat(quantityAggregated.toString()) || 0;
        if (reusableQtySum > 0 && quantityAggregatedF > 0) {
          reusablePct = Math.min((reusableQtySum / (quantityAggregatedF + reusableQtySum)) * 100, 100.0);
        }

        const [unitPriceForReuse] = getUnitPriceByPart(part, itemFinish, projectName);
        if (unitPriceForReuse !== null) {
          reusableCost = reusableQtySum * unitPriceForReuse * multiplier;
          totalReusableCost += reusableCost;
        }
      }

      // Calculate quantity display formats
      let quantityReqFt = 'N/A';
      let qtyStickReq = 'N/A';
      let quantityDisplayFormatted = `${quantityAggregated.toFixed(2)} ${displayUnit}`;

      if ((isProfile || isGasket) && part && part !== 'N/A') {
        quantityReqFt = `${quantityAggregated.toFixed(2)} ft`;
        const partData = (partsData as any)[part];
        const lengthStr = partData?.Length || '';
        const minPurchaseLength = parseLengthToFeet(lengthStr) || 1.0;
        if (minPurchaseLength > 0) {
          const numUnits = Math.ceil(quantityAggregated / minPurchaseLength);
          const unitLabel = isGasket ? 'rolls' : 'sticks';
          qtyStickReq = `${numUnits} (${minPurchaseLength.toFixed(0)}ft per)`;
        }

        // Format quantity display with breakdown
        const quantityList = item.quantity_list || [quantityAggregated];
        const validQuantities = quantityList.filter((q: any) => q !== null && q !== undefined && (typeof q === 'number' ? q > 0 : parseFloat(q.toString()) > 0));
        if (validQuantities.length > 0) {
          const lengthCounter: Record<string, number> = {};
          validQuantities.forEach((q: any) => {
            const val = typeof q === 'number' ? q : parseFloat(q.toString());
            const key = Math.round(val * 100) / 100;
            lengthCounter[key.toFixed(2)] = (lengthCounter[key.toFixed(2)] || 0) + 1;
          });
          const keys = Object.keys(lengthCounter).sort((a, b) => parseFloat(b) - parseFloat(a));
          if (keys.length > 1) {
            quantityDisplayFormatted = keys.map(k => {
              const count = lengthCounter[k];
              return count > 1 ? `${parseFloat(k).toFixed(0)}ft x${count}` : `${parseFloat(k).toFixed(0)}ft x1`;
            }).join(', ');
          } else if (keys.length === 1) {
            const lengthVal = parseFloat(keys[0]);
            const count = lengthCounter[keys[0]];
            quantityDisplayFormatted = count > 1 ? `${lengthVal.toFixed(0)}ft x${count}` : `${lengthVal.toFixed(0)}ft`;
          }
        }
      } else if (isAccessory && part && part !== 'N/A') {
        quantityReqFt = `${quantityAggregated.toFixed(2)} ${displayUnit}`;
        const partData = (partsData as any)[part];
        const unitsStr = partData?.Units || '1 pcs.';
        const lengthStr = partData?.Length || '';
        const lengthFt = parseLengthToFeet(lengthStr);
        let unitCountPerBundle = 1;
        let unitLabel = 'pcs per';
        if (lengthFt > 1.0) {
          unitCountPerBundle = lengthFt;
          unitLabel = 'ft per';
        } else {
          if (typeof unitsStr === 'string' && unitsStr.toLowerCase().includes('pc')) {
            const numPart = unitsStr.toLowerCase().split('pc')[0].trim();
            if (numPart) unitCountPerBundle = parseInt(numPart) || 1;
          }
        }
        qtyStickReq = `${unitCountPerBundle.toFixed(0)} ${unitLabel}`;
        const numOrders = Math.ceil(quantityAggregated / unitCountPerBundle);
        quantityDisplayFormatted = `${numOrders} order${numOrders !== 1 ? 's' : ''}`;
      } else {
        if (quantityAggregated > 0) {
          const unitPrice = originalTotalCostForItem / quantityAggregated;
          qtyStickReq = `$${unitPrice.toFixed(2)}`;
        } else {
          qtyStickReq = '$0.00';
        }
      }

      finalSummaryData.push({
        category,
        description,
        display,
        quantity_display: quantityDisplayFormatted,
        quantity_req_ft: quantityReqFt,
        qty_stick_req: qtyStickReq,
        original_total_cost: originalTotalCostForItem,
        total_cost: totalCostForItem,
        reusable_qty_display: reusableQtyDisplayString,
        reusable_pct: (isProfile || isGasket || isAccessory) ? reusablePct : 'N/A',
        reusable_cost: (isProfile || isGasket || isAccessory) ? reusableCost : 0.0,
        part,
        calculated_unit_type: calculatedUnitType
      });
    }
  }

  // Step 4: Write to worksheet with grouped sections
  let currentRow = 1;

  function getHeadersForCategory(category: string): string[] {
    if (category === 'PROFILES') {
      return [
        'Description', 'Project Total Materials', 'Total Feet', 'Sticks Required', 'Total Quantity Required',
        'Total List Cost', 'Discounted Total List Cost', 'Residual Material Quantity', 'Residual Waste %', 'Residual Material Cost'
      ];
    } else if (category === 'ACCESSORIES') {
      return [
        'Description', 'Project Total Materials', 'Total Pieces', 'Quantity Per Order', 'Orders Required',
        'Total List Cost', 'Discounted Total List Cost', 'Residual Material Quantity', 'Residual Waste %', 'Residual Material Cost'
      ];
    } else if (category === 'GASKETS') {
      return [
        'Description', 'Project Total Materials', 'Total Feet', 'Rolls Required', 'Total Quantity Required',
        'Total List Cost', 'Discounted Total List Cost', 'Residual Material Quantity', 'Residual Waste %', 'Residual Material Cost'
      ];
    } else {
      return [
        'Description', 'Project Total Materials', 'N/A', 'Unit Price', 'Total Quantity Required',
        'Total List Cost', 'Discounted Total List Cost', 'Residual Material Quantity', 'Residual Waste %', 'Residual Material Cost'
      ];
    }
  }

  for (const [category, items] of Object.entries(categories)) {
    if (!items || items.length === 0) continue;

    const headers = getHeadersForCategory(category);
    summarySheet.getCell(currentRow, 1).value = category;
    summarySheet.getCell(currentRow, 1).font = { bold: true, size: 12 };
    currentRow += 1;

    for (let col = 0; col < headers.length; col++) {
      const headerCell = summarySheet.getCell(currentRow, col + 1);
      headerCell.value = headers[col];
      headerCell.font = { bold: true };
      headerCell.border = { bottom: { style: 'thin' } };
    }
    currentRow += 1;

    let sectionOriginalTotal = 0.0;
    let sectionTotalCost = 0.0;
    let sectionResidualTotal = 0.0;

    for (const item of finalSummaryData) {
      if (item.category === category) {
        sectionOriginalTotal += item.original_total_cost;
        sectionTotalCost += item.total_cost;
        sectionResidualTotal += item.reusable_cost;

        summarySheet.getCell(currentRow, 1).value = item.description;
        summarySheet.getCell(currentRow, 2).value = item.display;
        summarySheet.getCell(currentRow, 3).value = item.quantity_req_ft;
        summarySheet.getCell(currentRow, 4).value = item.qty_stick_req;
        summarySheet.getCell(currentRow, 5).value = item.quantity_display;
        summarySheet.getCell(currentRow, 6).value = item.original_total_cost;
        summarySheet.getCell(currentRow, 6).numFmt = '$#,##0.00';
        summarySheet.getCell(currentRow, 7).value = item.total_cost;
        summarySheet.getCell(currentRow, 7).numFmt = '$#,##0.00';
        summarySheet.getCell(currentRow, 8).value = item.reusable_qty_display;
        summarySheet.getCell(currentRow, 9).value = typeof item.reusable_pct === 'number' ? `${item.reusable_pct.toFixed(2)}%` : item.reusable_pct;
        summarySheet.getCell(currentRow, 10).value = item.reusable_cost;
        summarySheet.getCell(currentRow, 10).numFmt = '$#,##0.00';
        currentRow += 1;
      }
    }

    grandOriginalTotal += sectionOriginalTotal;
    grandDiscountedTotal += sectionTotalCost;
    grandResidualTotal += sectionResidualTotal;

    // Add Section Totals
    const categoryMapping: Record<string, string> = {
      'PROFILES': 'Profile',
      'ACCESSORIES': 'Accessory',
      'GASKETS': 'Gasket',
      'DOORS': 'Door',
      'GLASS': 'Glass',
      'LABOR': 'Labor'
    };
    const categoryLabel = categoryMapping[category] || category;
    const totalLabel = `Total ${categoryLabel} Cost`;

    summarySheet.getCell(currentRow, 5).value = totalLabel;
    summarySheet.getCell(currentRow, 5).font = { bold: true };
    summarySheet.getCell(currentRow, 6).value = sectionOriginalTotal;
    summarySheet.getCell(currentRow, 6).numFmt = '$#,##0.00';
    summarySheet.getCell(currentRow, 6).font = { bold: true };
    summarySheet.getCell(currentRow, 7).value = sectionTotalCost;
    summarySheet.getCell(currentRow, 7).numFmt = '$#,##0.00';
    summarySheet.getCell(currentRow, 7).font = { bold: true };
    summarySheet.getCell(currentRow, 10).value = sectionResidualTotal;
    summarySheet.getCell(currentRow, 10).numFmt = '$#,##0.00';
    summarySheet.getCell(currentRow, 10).font = { bold: true };

    for (let col = 1; col <= 10; col++) {
      summarySheet.getCell(currentRow, col).border = { top: { style: 'thin' } };
    }
    currentRow += 2;
  }

  // Grand Totals Block
  const gtRow = currentRow + 2;

  summarySheet.getCell(gtRow, 6).value = 'Overall Total Price (List)';
  summarySheet.getCell(gtRow, 6).font = { bold: true };
  summarySheet.getCell(gtRow, 6).alignment = { horizontal: 'right' };
  summarySheet.getCell(gtRow, 6).border = { left: { style: 'thin' }, top: { style: 'thin' } };
  summarySheet.getCell(gtRow, 7).value = grandOriginalTotal;
  summarySheet.getCell(gtRow, 7).numFmt = '$#,##0.00';
  summarySheet.getCell(gtRow, 7).font = { bold: true };
  summarySheet.getCell(gtRow, 7).border = { right: { style: 'thin' }, top: { style: 'thin' } };

  summarySheet.getCell(gtRow + 1, 6).value = 'Overall Discounted Total';
  summarySheet.getCell(gtRow + 1, 6).font = { bold: true };
  summarySheet.getCell(gtRow + 1, 6).alignment = { horizontal: 'right' };
  summarySheet.getCell(gtRow + 1, 6).border = { left: { style: 'thin' } };
  summarySheet.getCell(gtRow + 1, 7).value = grandDiscountedTotal;
  summarySheet.getCell(gtRow + 1, 7).numFmt = '$#,##0.00';
  summarySheet.getCell(gtRow + 1, 7).font = { bold: true };
  summarySheet.getCell(gtRow + 1, 7).border = { right: { style: 'thin' } };

  summarySheet.getCell(gtRow + 2, 6).value = 'Overall Residual Cost';
  summarySheet.getCell(gtRow + 2, 6).font = { bold: true };
  summarySheet.getCell(gtRow + 2, 6).alignment = { horizontal: 'right' };
  summarySheet.getCell(gtRow + 2, 6).border = { left: { style: 'thin' } };
  summarySheet.getCell(gtRow + 2, 7).value = totalReusableCost;
  summarySheet.getCell(gtRow + 2, 7).numFmt = '$#,##0.00';
  summarySheet.getCell(gtRow + 2, 7).font = { bold: true };
  summarySheet.getCell(gtRow + 2, 7).border = { right: { style: 'thin' } };

  const reusePctOfGt = Math.min((totalReusableCost / grandDiscountedTotal * 100) || 0.0, 100.0);
  summarySheet.getCell(gtRow + 3, 6).value = 'Overall Waste %';
  summarySheet.getCell(gtRow + 3, 6).font = { bold: true };
  summarySheet.getCell(gtRow + 3, 6).alignment = { horizontal: 'right' };
  summarySheet.getCell(gtRow + 3, 6).border = { left: { style: 'thin' }, bottom: { style: 'thin' } };
  summarySheet.getCell(gtRow + 3, 7).value = `${reusePctOfGt.toFixed(2)}%`;
  summarySheet.getCell(gtRow + 3, 7).font = { bold: true };
  summarySheet.getCell(gtRow + 3, 7).border = { right: { style: 'thin' }, bottom: { style: 'thin' } };

  // Auto-fit columns
  for (let col = 1; col <= 10; col++) {
    let maxLength = 0;
    const lastRowNum = summarySheet.lastRow?.number || gtRow + 4;
    for (let row = 1; row <= lastRowNum; row++) {
      const cell = summarySheet.getCell(row, col);
      if (cell.value) {
        const cellValue = cell.value.toString();
        maxLength = Math.max(maxLength, cellValue.length);
      }
    }
    if (maxLength > 0) {
      summarySheet.getColumn(col).width = Math.min(maxLength + 2, 50);
    }
  }
}
