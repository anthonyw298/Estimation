export const SYSTEM_OPTIONS = [
  "YES 45TU FRONT SET(OG)",
  "YES 45TU CUSTOM",
  "Interior Storefront",
  "Other",
] as const

export const FINISH_OPTIONS = ["Clear", "Black", "Bronze", "Paint"] as const

export const DOOR_SIZES = [
  "None",
  "3' x 7'",
  "3' x 8'",
  "3' x 9'",
  "6' x 7'",
  "6' x 8'",
  "6' x 9'",
] as const

export const STILE_OPTIONS = ["Narrow", "Medium", "Wide"] as const

export const HARDWARE_OPTIONS = [
  "Continuous Hinges",
  "Concealed Closer",
  "Exit Devices",
  "Electric Strike",
  "Extended Ladder Pull (B2B)",
  "Extended Ladder Pull (Single)",
  "Latch Lock w/ Lever Handle",
  "Lever Handle",
] as const

export const FINISH_BADGE_MAP: Record<(typeof FINISH_OPTIONS)[number], string> =
  {
    Clear: "bg-sky-500/20 text-sky-700 dark:text-sky-200",
    Black: "bg-slate-500/30 text-slate-100",
    Bronze: "bg-amber-500/20 text-amber-900 dark:text-amber-100",
    Paint: "bg-rose-500/20 text-rose-900 dark:text-rose-100",
  }
