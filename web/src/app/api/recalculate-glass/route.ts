import { NextResponse } from 'next/server';
import { supabase } from '@/lib/database';
import { calculateYes45tuQuantities } from '@/lib/yes45tu';
import { calculateYes45tuCenterSetQuantities } from '@/lib/yes45tu-center-set';
import { calculate_door_info } from '@/lib/formulas';
import { getPriceByPart } from '@/lib/pricing';
import type { ElevationData, CalculatedOutput, DoorConfig, ProjectSettings } from '@/types';

const GASKET_PARTS = new Set(['E2-0052', 'E2-0053', 'E2-0065']);

function priceOutputs(
  rawOutputs: CalculatedOutput[],
  finish: string,
  glassPerSqft: number,
  fabCostPerJoint: number,
): CalculatedOutput[] {
  const priced: CalculatedOutput[] = [];
  for (const output of rawOutputs) {
    if (output.manual) {
      let totalPrice = 0;
      if (output.type === 'Glass' && typeof output.quantity === 'number') {
        totalPrice = output.quantity * (output.price ?? glassPerSqft);
      } else if (output.type === 'Fabrication' && typeof output.quantity === 'number') {
        totalPrice = output.quantity * (output.price ?? fabCostPerJoint);
      }
      priced.push({
        ...output,
        price: output.type === 'Calculations' ? undefined : totalPrice,
      });
    } else {
      const isGasket =
        (output.description?.toLowerCase().includes('gasket') ?? false) ||
        GASKET_PARTS.has(output.part_number);
      const isProfile = output.type === 'profiles';
      const useGroup = isProfile || isGasket;
      const [summaryPrice, unitType] = getPriceByPart(
        output.part_number,
        output.quantity,
        finish,
        null,
        true,
        useGroup,
        output.description,
      );
      priced.push({
        ...output,
        price: summaryPrice ?? 0,
        unit: unitType ?? undefined,
      });
    }
  }
  return priced;
}

/**
 * POST /api/recalculate-glass
 * Fully recalculates and re-prices all elevation outputs across all projects.
 */
export async function POST() {
  try {
    const { data: rows, error } = await supabase
      .from('elevations')
      .select('project_name, name, data');
    if (error) throw error;

    // Pre-fetch all doors
    const { data: allDoorRows } = await supabase.from('doors').select('project_name, elevation_name, data');
    const doorMap: Record<string, Record<string, DoorConfig[]>> = {};
    for (const dr of allDoorRows ?? []) {
      if (!doorMap[dr.project_name]) doorMap[dr.project_name] = {};
      doorMap[dr.project_name][dr.elevation_name] = Array.isArray(dr.data) ? dr.data : [];
    }

    // Pre-fetch all project settings
    const { data: allSettings } = await supabase.from('settings').select('project_name, data');
    const settingsMap: Record<string, ProjectSettings> = {};
    for (const s of allSettings ?? []) {
      settingsMap[s.project_name] = s.data as ProjectSettings;
    }

    let updated = 0;
    let skipped = 0;
    const errors: string[] = [];

    for (const row of rows ?? []) {
      const elev = row.data as ElevationData;
      const projectName: string = row.project_name;
      const elevName: string = row.name;

      if (!elev.opening_width_inches || !elev.opening_height_inches) {
        skipped++;
        continue;
      }

      try {
        const projSettings = settingsMap[projectName] ?? {};
        const systemType = elev.system_type || '';
        const baysWide = elev.bays_wide || 1;
        const baysTall = elev.bays_tall || 1;
        const totalCount = elev.total_count || 1;
        const openingWidth = elev.opening_width_inches;
        const openingHeight = elev.opening_height_inches;
        const finish = (elev.finish || 'clear').toLowerCase();
        const doors = doorMap[projectName]?.[elevName] ?? [];
        const doorOnly = elev.door_only || false;
        const glassPerSqft = elev.glass_per_sqft ?? (projSettings as ProjectSettings).glass_per_sqft ?? 10.5;
        const fabCostPerJoint = elev.fabrication_cost_per_joint ?? (projSettings as ProjectSettings).fabrication_cost_per_joint ?? 15;

        let newOutputs: CalculatedOutput[] = [];
        let singleElevOutputs: CalculatedOutput[] | undefined;

        if (doorOnly) {
          const doorItems = calculate_door_info(doors, finish, totalCount);
          newOutputs = doorItems.map(d => ({
            description: d.description,
            quantity: d.quantity,
            part_number: d.part_number,
            type: d.type,
            price: d.price * d.quantity,
            manual: true,
            hardware: d.hardware,
            Style: d.Style,
          }));
        } else {
          const calcFn = systemType === 'YES 45TU Center Set'
            ? calculateYes45tuCenterSetQuantities
            : calculateYes45tuQuantities;

          const rawOutputs = calcFn(
            baysWide, baysTall, totalCount, openingWidth, openingHeight, doors,
            elev.custom_bay_widths, elev.custom_bay_heights,
            glassPerSqft, fabCostPerJoint,
          );

          newOutputs = priceOutputs(rawOutputs, finish, glassPerSqft, fabCostPerJoint);

          // Door hardware items
          const doorItems = calculate_door_info(doors, finish, totalCount);
          for (const d of doorItems) {
            newOutputs.push({
              description: d.description,
              quantity: d.quantity,
              part_number: d.part_number,
              type: d.type,
              price: d.price * d.quantity,
              manual: true,
              hardware: d.hardware,
              Style: d.Style,
            });
          }

          // Single-elevation outputs (count=1) for per-elevation cost display
          if (totalCount > 1) {
            const singleRaw = calcFn(
              baysWide, baysTall, 1, openingWidth, openingHeight, doors,
              elev.custom_bay_widths, elev.custom_bay_heights,
              glassPerSqft, fabCostPerJoint,
            );
            singleElevOutputs = priceOutputs(singleRaw, finish, glassPerSqft, fabCostPerJoint);
            const singleDoorItems = calculate_door_info(doors, finish, 1);
            for (const d of singleDoorItems) {
              singleElevOutputs.push({
                description: d.description,
                quantity: d.quantity,
                part_number: d.part_number,
                type: d.type,
                price: d.price * d.quantity,
                manual: true,
                hardware: d.hardware,
                Style: d.Style,
              });
            }
          }
        }

        const updatedData: ElevationData = {
          ...elev,
          calculated_outputs: newOutputs,
          ...(singleElevOutputs ? { single_elevation_outputs: singleElevOutputs } : {}),
        };

        const { error: saveError } = await supabase
          .from('elevations')
          .update({ data: updatedData })
          .eq('project_name', projectName)
          .eq('name', elevName);

        if (saveError) throw saveError;
        updated++;
      } catch (e) {
        errors.push(`${projectName}/${elevName}: ${e instanceof Error ? e.message : String(e)}`);
      }
    }

    return NextResponse.json({
      success: true,
      updated,
      skipped,
      errors,
      message: `Recalculated ${updated} elevations across all projects.`,
    });
  } catch (e) {
    return NextResponse.json(
      { success: false, error: e instanceof Error ? e.message : String(e) },
      { status: 500 },
    );
  }
}
