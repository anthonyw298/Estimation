import { useEffect, useMemo, useState } from "react"
import {
  Download,
  PanelsTopLeft,
  Pencil,
  Plus,
  RotateCcw,
  Save,
  Trash2,
} from "lucide-react"

import { Badge } from "@/components/ui/badge"
import { Button } from "@/components/ui/button"
import {
  Card,
  CardContent,
  CardDescription,
  CardFooter,
  CardHeader,
  CardTitle,
} from "@/components/ui/card"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { ScrollArea } from "@/components/ui/scroll-area"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Separator } from "@/components/ui/separator"
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table"
import { Textarea } from "@/components/ui/textarea"

import { DOOR_SIZES, FINISH_BADGE_MAP, FINISH_OPTIONS, HARDWARE_OPTIONS, STILE_OPTIONS, SYSTEM_OPTIONS } from "../constants"
import type {
  ActionResult,
  Door,
  DoorSizeOption,
  Elevation,
  ElevationDraftInput,
  HardwareOption,
  Project,
  StileOption,
} from "../types"

const createId = () => {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID()
  }
  return `door-${Math.random().toString(36).slice(2, 10)}`
}

interface ElevationWorkspaceProps {
  project?: Project
  selectedElevation?: Elevation
  onSaveElevation: (projectId: string, payload: ElevationDraftInput) => ActionResult
  onDeleteElevation: (projectId: string, elevationId: string) => ActionResult
  onSelectElevation: (elevationId: string | undefined) => void
}

interface ElevationFormState {
  id?: string
  name: string
  system: string
  finish: string
  totalCount: string
  openingWidth: string
  openingHeight: string
  baysWide: string
  baysTall: string
  customBayWidths: string
  customBayHeights: string
  doors: Door[]
  notes: string
}

interface DoorFormState {
  id?: string
  size: DoorSizeOption
  count: string
  stile: StileOption
  hardware: HardwareOption[]
  notes: string
}

const DEFAULT_DRAFT: ElevationFormState = {
  name: "",
  system: SYSTEM_OPTIONS[0],
  finish: FINISH_OPTIONS[0],
  totalCount: "",
  openingWidth: "",
  openingHeight: "",
  baysWide: "",
  baysTall: "",
  customBayWidths: "",
  customBayHeights: "",
  doors: [],
  notes: "",
}

const DEFAULT_DOOR_FORM: DoorFormState = {
  size: DOOR_SIZES[0],
  count: "1",
  stile: STILE_OPTIONS[0],
  hardware: [],
  notes: "",
}

type FormErrors = Partial<Record<keyof ElevationFormState, string>>

const parseNumericList = (value: string) =>
  value
    .split(",")
    .map((item) => Number.parseFloat(item.trim()))
    .filter((num) => Number.isFinite(num) && num > 0)

const formatNumericList = (values: number[]) =>
  values.length ? values.map((value) => value.toString()).join(", ") : ""

export function ElevationWorkspace({
  project,
  selectedElevation,
  onSaveElevation,
  onDeleteElevation,
  onSelectElevation,
}: ElevationWorkspaceProps) {
  const [draft, setDraft] = useState<ElevationFormState>(DEFAULT_DRAFT)
  const [doorForm, setDoorForm] = useState<DoorFormState>(DEFAULT_DOOR_FORM)
  const [formErrors, setFormErrors] = useState<FormErrors>({})
  const [doorError, setDoorError] = useState<string | null>(null)
  const [isSaving, setIsSaving] = useState(false)

  const isEditingDoor = Boolean(doorForm.id)

  /* eslint-disable react-hooks/set-state-in-effect */
  useEffect(() => {
    if (!selectedElevation) {
      setDraft((prev) => ({
        ...DEFAULT_DRAFT,
        system: prev.system,
        finish: prev.finish,
      }))
      setDoorForm(DEFAULT_DOOR_FORM)
      setFormErrors({})
      return
    }
    setDraft({
      id: selectedElevation.id,
      name: selectedElevation.name,
      system: selectedElevation.system,
      finish: selectedElevation.finish,
      totalCount: selectedElevation.totalCount.toString(),
      openingWidth: selectedElevation.openingWidth.toString(),
      openingHeight: selectedElevation.openingHeight.toString(),
      baysWide: selectedElevation.baysWide?.toString() ?? "",
      baysTall: selectedElevation.baysTall?.toString() ?? "",
      customBayWidths: formatNumericList(selectedElevation.customBayWidths),
      customBayHeights: formatNumericList(selectedElevation.customBayHeights),
      doors: selectedElevation.doors,
      notes: selectedElevation.notes ?? "",
    })
    setDoorForm(DEFAULT_DOOR_FORM)
    setFormErrors({})
  }, [selectedElevation])
  /* eslint-enable react-hooks/set-state-in-effect */

  const handleFieldChange =
    (field: keyof ElevationFormState) => (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
      setDraft((prev) => ({ ...prev, [field]: event.target.value }))
      setFormErrors((prev) => ({ ...prev, [field]: undefined }))
    }

  const handleSelectChange =
    (field: keyof ElevationFormState) => (value: string) => {
      setDraft((prev) => ({ ...prev, [field]: value }))
      setFormErrors((prev) => ({ ...prev, [field]: undefined }))
    }

  const validateDraft = (): boolean => {
    const errors: FormErrors = {}
    if (!draft.name.trim()) {
      errors.name = "Elevation name is required."
    }
    if (!draft.totalCount || Number.isNaN(Number(draft.totalCount))) {
      errors.totalCount = "Provide the total count."
    }
    if (!draft.openingWidth || Number.isNaN(Number(draft.openingWidth))) {
      errors.openingWidth = "Provide the opening width."
    }
    if (!draft.openingHeight || Number.isNaN(Number(draft.openingHeight))) {
      errors.openingHeight = "Provide the opening height."
    }

    setFormErrors(errors)
    return !Object.keys(errors).length
  }

  const buildPayload = (): ElevationDraftInput | null => {
    if (!validateDraft()) {
      return null
    }
    const totalCount = Number.parseInt(draft.totalCount, 10)
    const openingWidth = Number.parseFloat(draft.openingWidth)
    const openingHeight = Number.parseFloat(draft.openingHeight)
    const baysWide = draft.baysWide ? Number.parseInt(draft.baysWide, 10) : undefined
    const baysTall = draft.baysTall ? Number.parseInt(draft.baysTall, 10) : undefined

    return {
      id: draft.id,
      name: draft.name.trim(),
      system: draft.system as ElevationDraftInput["system"],
      finish: draft.finish as ElevationDraftInput["finish"],
      totalCount,
      openingWidth,
      openingHeight,
      baysWide: Number.isFinite(baysWide ?? NaN) ? baysWide : undefined,
      baysTall: Number.isFinite(baysTall ?? NaN) ? baysTall : undefined,
      customBayWidths: parseNumericList(draft.customBayWidths),
      customBayHeights: parseNumericList(draft.customBayHeights),
      doors: draft.doors,
      notes: draft.notes.trim() || undefined,
    }
  }

  const handleSave = () => {
    if (!project) return
    const payload = buildPayload()
    if (!payload) return
    setIsSaving(true)
    onSaveElevation(project.id, payload)
    setIsSaving(false)
  }

  const handleDelete = () => {
    if (!project || !draft.id) return
    if (window.confirm(`Delete elevation “${draft.name}”? This cannot be undone.`)) {
      onDeleteElevation(project.id, draft.id)
      setDraft(DEFAULT_DRAFT)
      setDoorForm(DEFAULT_DOOR_FORM)
    }
  }

  const handleDoorField =
    (field: keyof DoorFormState) =>
    (
      value:
        | string
        | React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>
        | React.ChangeEvent<HTMLTextAreaElement>
    ) => {
      const nextValue = typeof value === "string" ? value : value.target.value
      setDoorForm((prev) => ({ ...prev, [field]: nextValue }))
      setDoorError(null)
    }

  const toggleHardware = (option: HardwareOption) => {
    setDoorForm((prev) => {
      const exists = prev.hardware.includes(option)
      return {
        ...prev,
        hardware: exists
          ? prev.hardware.filter((item) => item !== option)
          : [...prev.hardware, option],
      }
    })
  }

  const handleDoorSubmit = () => {
    if (doorForm.size === DOOR_SIZES[0]) {
      setDoorError("Select a door size.")
      return
    }
    const count = Number.parseInt(doorForm.count, 10)
    if (!Number.isFinite(count) || count <= 0) {
      setDoorError("Provide a valid door quantity.")
      return
    }
    const nextDoor: Door = {
      id: doorForm.id ?? createId(),
      size: doorForm.size,
      count,
      stile: doorForm.stile,
      hardware: doorForm.hardware,
      notes: doorForm.notes.trim() || undefined,
    }

    setDraft((prev) => ({
      ...prev,
      doors: doorForm.id
        ? prev.doors.map((door) => (door.id === doorForm.id ? nextDoor : door))
        : [...prev.doors, nextDoor],
    }))
    setDoorForm(DEFAULT_DOOR_FORM)
    setDoorError(null)
  }

  const handleDoorEdit = (door: Door) => {
    setDoorForm({
      id: door.id,
      size: door.size,
      count: door.count.toString(),
      stile: door.stile,
      hardware: door.hardware,
      notes: door.notes ?? "",
    })
  }

  const handleDoorDelete = (doorId: string) => {
    setDraft((prev) => ({
      ...prev,
      doors: prev.doors.filter((door) => door.id !== doorId),
    }))
    if (doorForm.id === doorId) {
      setDoorForm(DEFAULT_DOOR_FORM)
    }
  }

  const handleDoorReset = () => {
    setDoorForm(DEFAULT_DOOR_FORM)
    setDoorError(null)
  }

  const metrics = useMemo(() => {
    const totalCount = Number.parseInt(draft.totalCount, 10) || 0
    const widthInches = Number.parseFloat(draft.openingWidth) || 0
    const heightInches = Number.parseFloat(draft.openingHeight) || 0
    const areaSqFt = ((widthInches * heightInches) / 144) * totalCount
    const perimeterFt =
      ((widthInches * 2 + heightInches * 2) / 12) * totalCount
    const totalDoors = draft.doors.reduce((sum, door) => sum + door.count, 0)
    const hardwareBreakdown = draft.doors.reduce<Record<string, number>>(
      (acc, door) => {
        door.hardware.forEach((item) => {
          acc[item] = (acc[item] ?? 0) + door.count
        })
        return acc
      },
      {}
    )
    return {
      areaSqFt: Number.isFinite(areaSqFt) ? areaSqFt : 0,
      perimeterFt: Number.isFinite(perimeterFt) ? perimeterFt : 0,
      totalDoors,
      hardwareBreakdown,
    }
  }, [draft.doors, draft.openingHeight, draft.openingWidth, draft.totalCount])

  if (!project) {
    return (
      <Card className="border-dashed">
        <CardHeader>
          <CardTitle>Pick a project to continue</CardTitle>
          <CardDescription>
            Create or select a project to unlock elevation planning tools.
          </CardDescription>
        </CardHeader>
      </Card>
    )
  }

  return (
    <div className="space-y-6">
      <Card>
        <CardHeader className="flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
          <div className="space-y-2">
            <CardTitle>Elevation workspace</CardTitle>
            <CardDescription>
              Configure geometry, finish selections, and paired door packages.
            </CardDescription>
          </div>
          <div className="flex flex-wrap items-center gap-4">
            <div className="space-y-2">
              <Label>Saved elevations</Label>
              <Select
                value={selectedElevation?.id ?? "__new"}
                onValueChange={(value) =>
                  value === "__new"
                    ? onSelectElevation(undefined)
                    : onSelectElevation(value)
                }
              >
                <SelectTrigger className="w-[220px]">
                  <SelectValue placeholder="Choose elevation" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="__new">＋ Start fresh</SelectItem>
                  {project.elevations.map((elevation) => (
                    <SelectItem key={elevation.id} value={elevation.id}>
                      {elevation.name}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            {selectedElevation ? (
              <Badge className={FINISH_BADGE_MAP[selectedElevation.finish]}>
                Finish: {selectedElevation.finish}
              </Badge>
            ) : null}
          </div>
        </CardHeader>
      </Card>

      <div className="grid gap-6 lg:grid-cols-[1.7fr,1fr]">
        <div className="space-y-6">
          <Card>
            <CardHeader>
              <CardTitle>General details</CardTitle>
              <CardDescription>System and quantity level settings.</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-2">
                  <Label htmlFor="elevation-name">Elevation name</Label>
                  <Input
                    id="elevation-name"
                    value={draft.name}
                    onChange={handleFieldChange("name")}
                    placeholder="Lobby Entry A"
                  />
                  {formErrors.name ? (
                    <p className="text-sm text-destructive">{formErrors.name}</p>
                  ) : null}
                </div>
                <div className="space-y-2">
                  <Label>System</Label>
                  <Select
                    value={draft.system}
                    onValueChange={handleSelectChange("system")}
                  >
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {SYSTEM_OPTIONS.map((option) => (
                        <SelectItem key={option} value={option}>
                          {option}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>
              <div className="grid gap-4 md:grid-cols-3">
                <div className="space-y-1.5">
                  <Label>Total count</Label>
                  <Input
                    value={draft.totalCount}
                    onChange={handleFieldChange("totalCount")}
                    placeholder="2"
                    type="number"
                    min={0}
                  />
                  {formErrors.totalCount ? (
                    <p className="text-sm text-destructive">
                      {formErrors.totalCount}
                    </p>
                  ) : null}
                </div>
                <div className="space-y-1.5">
                  <Label>Finish</Label>
                  <Select
                    value={draft.finish}
                    onValueChange={handleSelectChange("finish")}
                  >
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {FINISH_OPTIONS.map((option) => (
                        <SelectItem key={option} value={option}>
                          {option}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-1.5">
                  <Label>Notes</Label>
                  <Input
                    value={draft.notes}
                    onChange={handleFieldChange("notes")}
                    placeholder="Optional remarks"
                  />
                </div>
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Geometry + custom bays</CardTitle>
              <CardDescription>Capture opening sizes and bespoke bay layouts.</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-1.5">
                  <Label>Opening width (in)</Label>
                  <Input
                    value={draft.openingWidth}
                    onChange={handleFieldChange("openingWidth")}
                    placeholder="144"
                  />
                  {formErrors.openingWidth ? (
                    <p className="text-sm text-destructive">
                      {formErrors.openingWidth}
                    </p>
                  ) : null}
                </div>
                <div className="space-y-1.5">
                  <Label>Opening height (in)</Label>
                  <Input
                    value={draft.openingHeight}
                    onChange={handleFieldChange("openingHeight")}
                    placeholder="120"
                  />
                  {formErrors.openingHeight ? (
                    <p className="text-sm text-destructive">
                      {formErrors.openingHeight}
                    </p>
                  ) : null}
                </div>
              </div>
              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-1.5">
                  <Label># Bays wide</Label>
                  <Input
                    value={draft.baysWide}
                    onChange={handleFieldChange("baysWide")}
                    placeholder="3"
                  />
                </div>
                <div className="space-y-1.5">
                  <Label># Bays tall</Label>
                  <Input
                    value={draft.baysTall}
                    onChange={handleFieldChange("baysTall")}
                    placeholder="2"
                  />
                </div>
              </div>
              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-1.5">
                  <Label>Custom bay widths (in)</Label>
                  <Textarea
                    value={draft.customBayWidths}
                    onChange={handleFieldChange("customBayWidths")}
                    placeholder="eg. 36, 48, 36"
                    rows={2}
                  />
                </div>
                <div className="space-y-1.5">
                  <Label>Custom bay heights (in)</Label>
                  <Textarea
                    value={draft.customBayHeights}
                    onChange={handleFieldChange("customBayHeights")}
                    placeholder="eg. 60, 60"
                    rows={2}
                  />
                </div>
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <CardTitle>Door packages</CardTitle>
                <CardDescription>Track each stile combination and hardware pairing.</CardDescription>
              </div>
              {doorError ? (
                <p className="text-sm text-destructive">{doorError}</p>
              ) : null}
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-1.5">
                  <Label>Door size</Label>
                  <Select
                    value={doorForm.size}
                    onValueChange={(value) =>
                      setDoorForm((prev) => ({ ...prev, size: value as DoorSizeOption }))
                    }
                  >
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {DOOR_SIZES.map((size) => (
                        <SelectItem key={size} value={size}>
                          {size}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-1.5">
                  <Label>Quantity</Label>
                  <Input
                    type="number"
                    min={1}
                    value={doorForm.count}
                    onChange={(event) =>
                      setDoorForm((prev) => ({ ...prev, count: event.target.value }))
                    }
                  />
                </div>
              </div>
              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-1.5">
                  <Label>Stile</Label>
                  <Select
                    value={doorForm.stile}
                    onValueChange={(value) =>
                      setDoorForm((prev) => ({ ...prev, stile: value as StileOption }))
                    }
                  >
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {STILE_OPTIONS.map((option) => (
                        <SelectItem key={option} value={option}>
                          {option}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-1.5">
                  <Label>Notes</Label>
                  <Input
                    value={doorForm.notes}
                    onChange={handleDoorField("notes")}
                    placeholder="Optional hardware notes"
                  />
                </div>
              </div>
              <div className="space-y-2">
                <Label>Hardware</Label>
                <ScrollArea className="h-[120px] rounded-md border p-3">
                  <div className="flex flex-wrap gap-2">
                    {HARDWARE_OPTIONS.map((option) => {
                      const active = doorForm.hardware.includes(option)
                      return (
                        <Button
                          key={option}
                          type="button"
                          size="sm"
                          variant={active ? "default" : "outline"}
                          onClick={() => toggleHardware(option)}
                        >
                          {option}
                        </Button>
                      )
                    })}
                  </div>
                </ScrollArea>
              </div>
              <div className="flex flex-wrap gap-3">
                <Button onClick={handleDoorSubmit} className="gap-2">
                  {isEditingDoor ? (
                    <>
                      <Save className="h-4 w-4" />
                      Update door
                    </>
                  ) : (
                    <>
                      <Plus className="h-4 w-4" />
                      Add door
                    </>
                  )}
                </Button>
                {isEditingDoor ? (
                  <Button
                    type="button"
                    variant="ghost"
                    className="gap-2"
                    onClick={handleDoorReset}
                  >
                    <RotateCcw className="h-4 w-4" />
                    Reset
                  </Button>
                ) : null}
              </div>
              <Separator />
              <div className="space-y-3">
                <Label className="text-sm">Current doors</Label>
                <ScrollArea className="h-[220px] rounded-md border">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>Size</TableHead>
                        <TableHead>Count</TableHead>
                        <TableHead>Stile</TableHead>
                        <TableHead>Hardware</TableHead>
                        <TableHead className="w-[110px] text-right">Actions</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {draft.doors.map((door) => (
                        <TableRow key={door.id}>
                          <TableCell>{door.size}</TableCell>
                          <TableCell>{door.count}</TableCell>
                          <TableCell>{door.stile}</TableCell>
                          <TableCell>
                            <div className="flex flex-wrap gap-1">
                              {door.hardware.map((item) => (
                                <Badge key={item} variant="outline">
                                  {item}
                                </Badge>
                              ))}
                            </div>
                          </TableCell>
                          <TableCell className="text-right">
                            <div className="flex justify-end gap-2">
                              <Button
                                size="icon"
                                variant="ghost"
                                onClick={() => handleDoorEdit(door)}
                                aria-label="Edit door"
                              >
                                <Pencil className="h-4 w-4" />
                              </Button>
                              <Button
                                size="icon"
                                variant="ghost"
                                className="text-destructive"
                                onClick={() => handleDoorDelete(door.id)}
                                aria-label="Delete door"
                              >
                                <Trash2 className="h-4 w-4" />
                              </Button>
                            </div>
                          </TableCell>
                        </TableRow>
                      ))}
                      {!draft.doors.length ? (
                        <TableRow>
                          <TableCell colSpan={5} className="text-center text-muted-foreground">
                            No doors yet — add a package to get started.
                          </TableCell>
                        </TableRow>
                      ) : null}
                    </TableBody>
                  </Table>
                </ScrollArea>
              </div>
            </CardContent>
          </Card>
        </div>

        <div className="space-y-6">
          <Card>
            <CardHeader>
              <CardTitle>Quick metrics</CardTitle>
              <CardDescription>
                Real-time calculations to sense-check your takeoff.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-2 gap-3">
                <MetricTile
                  label="Glass area"
                  value={`${metrics.areaSqFt.toFixed(1)} ft²`}
                  detail="based on width × height × count"
                />
                <MetricTile
                  label="Perimeter"
                  value={`${metrics.perimeterFt.toFixed(1)} ft`}
                  detail="converted from inches"
                />
                <MetricTile
                  label="Door quantity"
                  value={metrics.totalDoors.toString()}
                  detail="total leaf count"
                />
                <MetricTile
                  label="Hardware styles"
                  value={Object.keys(metrics.hardwareBreakdown).length.toString()}
                  detail="unique hardware lines"
                />
              </div>
              <Separator />
              <div>
                <p className="mb-2 text-sm font-medium">Hardware coverage</p>
                {Object.keys(metrics.hardwareBreakdown).length ? (
                  <ul className="space-y-1 text-sm text-muted-foreground">
                    {Object.entries(metrics.hardwareBreakdown).map(([item, qty]) => (
                      <li key={item} className="flex items-center justify-between rounded-md bg-muted/40 px-3 py-1.5">
                        <span>{item}</span>
                        <span className="font-semibold text-foreground">{qty}</span>
                      </li>
                    ))}
                  </ul>
                ) : (
                  <p className="text-sm text-muted-foreground">
                    Add doors to see a breakdown per hardware type.
                  </p>
                )}
              </div>
            </CardContent>
            <CardFooter className="flex flex-wrap gap-3">
              <Button className="gap-2" onClick={handleSave} disabled={isSaving}>
                <Save className="h-4 w-4" />
                Save elevation
              </Button>
              <Button
                variant="outline"
                className="gap-2"
                onClick={() => alert("Excel export coming soon.")}
              >
                <Download className="h-4 w-4" />
                Generate report
              </Button>
              {draft.id ? (
                <Button
                  variant="ghost"
                  className="gap-2 text-destructive hover:text-destructive"
                  onClick={handleDelete}
                >
                  <Trash2 className="h-4 w-4" />
                  Delete
                </Button>
              ) : null}
            </CardFooter>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Elevation digest</CardTitle>
              <CardDescription>A snapshot of the active elevation.</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4 text-sm">
              <div className="space-y-1">
                <p className="text-xs uppercase text-muted-foreground">System</p>
                <p className="text-base font-semibold">{draft.system}</p>
              </div>
              <div className="space-y-1">
                <p className="text-xs uppercase text-muted-foreground">Finish</p>
                <p className="text-base font-semibold">{draft.finish}</p>
              </div>
              <div className="space-y-1">
                <p className="text-xs uppercase text-muted-foreground">Custom bays</p>
                <p>
                  {draft.customBayWidths
                    ? `${draft.customBayWidths} in wide`
                    : "Uniform widths"}
                </p>
                <p>
                  {draft.customBayHeights
                    ? `${draft.customBayHeights} in tall`
                    : "Uniform heights"}
                </p>
              </div>
              <Separator />
              <div className="flex items-center gap-3 rounded-lg border bg-muted/40 p-3">
                <PanelsTopLeft className="h-8 w-8 text-primary" />
                <div>
                  <p className="font-semibold">Preview</p>
                  <p className="text-sm text-muted-foreground">
                    Visual explorer coming soon — for now this card summarizes your setup.
                  </p>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  )
}

function MetricTile({
  label,
  value,
  detail,
}: {
  label: string
  value: string
  detail: string
}) {
  return (
    <div className="rounded-xl border bg-card/60 p-3">
      <p className="text-xs uppercase tracking-wide text-muted-foreground">{label}</p>
      <p className="text-2xl font-semibold">{value}</p>
      <p className="text-xs text-muted-foreground">{detail}</p>
    </div>
  )
}
