// Elevation data stored in database
export interface ElevationData {
  system_type: string;
  finish: string; // 'Clear' | 'Black' | 'Paint'
  opening_width_inches: number;
  opening_height_inches: number;
  bays_wide: number;
  bays_tall: number;
  total_count: number;
  custom_bay_widths?: number[];
  custom_bay_heights?: number[];
  glass_per_sqft?: number;
  fabrication_cost_per_joint?: number;
  calculated_outputs?: CalculatedOutput[];
  material_impacts?: MaterialImpactDetails[];
  door_only?: boolean;
  sqft_per_type?: Record<string, number>;
  total_sqft?: number;
  perimeter_ft?: number;
  total_perimeter_ft?: number;
}

// A single calculated output item (profile, accessory, glass, fabrication, door)
export interface CalculatedOutput {
  description: string;
  quantity: number | number[];
  part_number: string;
  type: string; // 'profiles' | 'accessories' | 'Glass' | 'Fabrication' | 'Doors' | 'Calculations'
  price?: number;
  unit?: string;
  manual?: boolean;
  message?: string;
  hardware?: Record<string, boolean>;
  Style?: string;
}

// Door configuration
export interface DoorConfig {
  size: string; // e.g. "3' X 7'"
  count: number;
  stile: string; // 'Narrow' | 'Medium' | 'Wide'
  hardware: Record<string, boolean>;
  x_in?: number;
  x_positions?: number[];
}

// Material impact tracking for inventory management
export interface MaterialImpactDetails {
  part_number: string;
  requested_qty: number | number[];
  purchased_qty_or_length: number;
  leftover_generated_qty_or_length: number;
  used_from_leftover_qty_or_length: number;
  cost_incurred: number;
  type_processed_as: 'profile' | 'accessory' | null;
  finish?: string;
  description?: string;
  is_bay_width_list?: boolean;
  bay_widths_processed?: number[];
  all_new_leftovers?: number[];
  leftover_pieces_consumed?:
    | Array<{ original_length: number; used_length: number }>
    | Array<[number, number]>;
}

// Extra materials / leftover inventory
export interface ExtraMaterial {
  quantity: number;
  length_pieces: number[];
}

// Project settings
export interface ProjectSettings {
  // Pricing Adjustment tab
  discount_multiplier?: number;
  discount_multiplier_low?: number;
  discount_multiplier_high?: number;
  discount_threshold?: number;
  glass_per_sqft?: number;
  fabrication_cost_per_joint?: number;

  // Additional Cost Settings (Summary tab)
  overhead_materials_pct?: number;
  overhead_labor_pct?: number;
  admin_management_pct?: number;
  engineering_pct?: number;
  packaging_materials_pct?: number;
  shipping_transport_pct?: number;
  commissions_pct?: number;

  // Markup Settings (Summary tab)
  profit_on_material_pct?: number;
  profit_on_waste_pct?: number;
  profit_on_glass_pct?: number;
  profit_on_wages_pct?: number;
  planning_technical_pct?: number;
  commission_pct?: number;

  // Elevation Summary Display checkboxes
  show_elevation_names?: boolean;
  show_elevation_quantity?: boolean;
  show_elevation_dimensions?: boolean;
  show_elevation_sqft?: boolean;
  show_elevation_perimeter?: boolean;

  // Legacy
  additional_costs?: AdditionalCost[];
}

export interface AdditionalCost {
  description: string;
  amount: number;
  type: 'fixed' | 'percentage';
}

// Report configuration for export customization
export interface ReportConfig {
  elevations_included: Record<string, boolean>;
  summary_included: boolean;
  per_elevation_sections: Record<string, Record<string, boolean>>;
  per_elevation_columns: Record<string, Record<string, Record<string, boolean>>>;
  summary_options: {
    sections: Record<string, boolean>;
    columns: Record<string, Record<string, boolean>>;
    cost_overview: Record<string, boolean>;
  };
}

// Waste analysis types
export interface WasteStatistics {
  total_waste_cost: number;
  total_material_cost: number;
  overall_waste_percentage: number;
  material_breakdown: WasteMaterialBreakdown[];
  suggestions: WasteSuggestion[];
}

export interface WasteMaterialBreakdown {
  part_number: string;
  description: string;
  finish: string;
  total_quantity: number;
  waste_quantity: number;
  waste_quantity_display: string;
  waste_percentage: number;
  waste_cost: number;
  material_cost: number;
  unit: string;
  individual_pieces: number[];
}

export interface WasteSuggestion {
  priority: 'high' | 'medium' | 'low';
  category: string;
  message: string;
  estimated_savings: number | null;
}

// Door pricing types
export interface DoorPriceMatrix {
  [sizeKey: string]: {
    [stileType: string]: {
      [finish: string]: number;
    };
  };
}

export interface HardwarePrices {
  [hardware: string]: {
    [finish: string]: number;
  };
}

// Report data
export interface ReportData {
  project_name: string;
  elevations: Record<string, ElevationData>;
  settings: ProjectSettings;
  materials: Record<string, ExtraMaterial>;
  doors: Record<string, DoorConfig[]>;
  total_cost: number;
  discount_multiplier: number;
  discounted_total: number;
  waste_cost: number;
  waste_percentage: number;
}
