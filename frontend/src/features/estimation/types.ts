import type {
  DOOR_SIZES,
  FINISH_OPTIONS,
  HARDWARE_OPTIONS,
  STILE_OPTIONS,
  SYSTEM_OPTIONS,
} from "./constants"

export type SystemOption = (typeof SYSTEM_OPTIONS)[number]
export type FinishOption = (typeof FINISH_OPTIONS)[number]
export type DoorSizeOption = (typeof DOOR_SIZES)[number]
export type StileOption = (typeof STILE_OPTIONS)[number]
export type HardwareOption = (typeof HARDWARE_OPTIONS)[number]

export type StatusTone = "success" | "error" | "info"

export interface StatusMessage {
  id: string
  tone: StatusTone
  text: string
  timestamp: number
}

export interface Door {
  id: string
  size: DoorSizeOption
  count: number
  stile: StileOption
  hardware: HardwareOption[]
  notes?: string
}

export interface Elevation {
  id: string
  name: string
  system: SystemOption
  finish: FinishOption
  totalCount: number
  openingWidth: number
  openingHeight: number
  baysWide?: number | null
  baysTall?: number | null
  customBayWidths: number[]
  customBayHeights: number[]
  doors: Door[]
  notes?: string
  createdAt: string
  updatedAt: string
}

export interface Project {
  id: string
  name: string
  createdAt: string
  updatedAt: string
  elevations: Elevation[]
}

export interface EstimationState {
  projects: Project[]
  selectedProjectId?: string
  selectedElevationId?: string
  status?: StatusMessage
}

export interface ElevationDraftInput {
  id?: string
  name: string
  system: SystemOption
  finish: FinishOption
  totalCount: number
  openingWidth: number
  openingHeight: number
  baysWide?: number | null
  baysTall?: number | null
  customBayWidths: number[]
  customBayHeights: number[]
  doors: Door[]
  notes?: string
}

export interface ActionResult {
  ok: boolean
  message: string
}
