import { useCallback, useEffect, useMemo, useState } from "react"

import { DOOR_SIZES, FINISH_OPTIONS, STILE_OPTIONS, SYSTEM_OPTIONS } from "./constants"
import type {
  ActionResult,
  Elevation,
  ElevationDraftInput,
  EstimationState,
  Project,
  StatusMessage,
} from "./types"

const STORAGE_KEY = "estimation:workspace:v1"

const createId = () => {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID()
  }
  return `id-${Math.random().toString(36).slice(2, 10)}`
}

const seedState = (): EstimationState => {
  const now = new Date().toISOString()
  const projectId = createId()
  const elevationId = createId()
  return {
    projects: [
      {
        id: projectId,
        name: "Flagship Lobby",
        createdAt: now,
        updatedAt: now,
        elevations: [
          {
            id: elevationId,
            name: "Front Entry",
            system: SYSTEM_OPTIONS[0],
            finish: FINISH_OPTIONS[0],
            totalCount: 2,
            openingWidth: 144,
            openingHeight: 120,
            baysWide: 3,
            baysTall: 2,
            customBayWidths: [],
            customBayHeights: [],
            notes: "Demo elevation to showcase the workflow.",
            doors: [
              {
                id: createId(),
                size: DOOR_SIZES[1],
                count: 2,
                stile: STILE_OPTIONS[1],
                hardware: [
                  "Continuous Hinges",
                  "Concealed Closer",
                  "Lever Handle",
                ],
              },
            ],
            createdAt: now,
            updatedAt: now,
          },
        ],
      },
    ],
    selectedProjectId: projectId,
    selectedElevationId: elevationId,
    status: undefined,
  }
}

const hydrateState = (): EstimationState => {
  if (typeof window === "undefined") {
    return seedState()
  }
  try {
    const stored = window.localStorage.getItem(STORAGE_KEY)
    if (!stored) return seedState()
    const parsed = JSON.parse(stored) as EstimationState
    return {
      ...parsed,
      projects: parsed.projects ?? [],
    }
  } catch {
    return seedState()
  }
}

const makeStatus = (tone: StatusMessage["tone"], text: string): StatusMessage => ({
  id: createId(),
  tone,
  text,
  timestamp: Date.now(),
})

export function useEstimationState() {
  const [state, setState] = useState<EstimationState>(() => hydrateState())

  useEffect(() => {
    if (typeof window === "undefined") return
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state))
  }, [state])

  const createProject = useCallback(
    (name: string): ActionResult => {
      const trimmed = name.trim()
      let result: ActionResult = { ok: false, message: "" }
      setState((prev) => {
        if (!trimmed) {
          result = { ok: false, message: "Project name is required." }
          return { ...prev, status: makeStatus("error", result.message) }
        }
        if (
          prev.projects.some(
            (proj) => proj.name.toLowerCase() === trimmed.toLowerCase()
          )
        ) {
          result = { ok: false, message: "A project with that name already exists." }
          return { ...prev, status: makeStatus("error", result.message) }
        }
        const now = new Date().toISOString()
        const newProject: Project = {
          id: createId(),
          name: trimmed,
          createdAt: now,
          updatedAt: now,
          elevations: [],
        }
        result = { ok: true, message: `Project “${trimmed}” created.` }
        return {
          ...prev,
          projects: [newProject, ...prev.projects],
          selectedProjectId: newProject.id,
          selectedElevationId: undefined,
          status: makeStatus("success", result.message),
        }
      })
      return result
    },
    []
  )

  const deleteProject = useCallback(
    (projectId: string): ActionResult => {
      let result: ActionResult = { ok: false, message: "" }
      setState((prev) => {
        const target = prev.projects.find((proj) => proj.id === projectId)
        if (!target) {
          result = { ok: false, message: "Project not found." }
          return { ...prev, status: makeStatus("error", result.message) }
        }
        const remaining = prev.projects.filter((proj) => proj.id !== projectId)
        const nextSelectedProjectId =
          prev.selectedProjectId === projectId
            ? remaining[0]?.id
            : prev.selectedProjectId
        const nextSelectedElevationId =
          prev.selectedProjectId === projectId
            ? remaining[0]?.elevations[0]?.id
            : prev.selectedElevationId
        result = { ok: true, message: `Project “${target.name}” deleted.` }
        return {
          ...prev,
          projects: remaining,
          selectedProjectId: nextSelectedProjectId,
          selectedElevationId: nextSelectedElevationId,
          status: makeStatus("success", result.message),
        }
      })
      return result
    },
    []
  )

  const selectProject = useCallback((projectId: string | undefined) => {
    setState((prev) => {
      const project = prev.projects.find((proj) => proj.id === projectId)
      return {
        ...prev,
        selectedProjectId: project?.id,
        selectedElevationId: project?.elevations[0]?.id,
      }
    })
  }, [])

  const selectElevation = useCallback((elevationId: string | undefined) => {
    setState((prev) => ({
      ...prev,
      selectedElevationId: elevationId,
    }))
  }, [])

  const saveElevation = useCallback(
    (projectId: string, payload: ElevationDraftInput): ActionResult => {
      let result: ActionResult = { ok: false, message: "" }
      setState((prev) => {
        const projectIndex = prev.projects.findIndex(
          (proj) => proj.id === projectId
        )
        if (projectIndex === -1) {
          result = {
            ok: false,
            message: "Please select a project before saving an elevation.",
          }
          return { ...prev, status: makeStatus("error", result.message) }
        }
        const project = prev.projects[projectIndex]
        const now = new Date().toISOString()
        const elevationId = payload.id ?? createId()
        const existing = project.elevations.find((el) => el.id === elevationId)
        const elevation: Elevation = {
          id: elevationId,
          name: payload.name,
          system: payload.system,
          finish: payload.finish,
          totalCount: payload.totalCount,
          openingWidth: payload.openingWidth,
          openingHeight: payload.openingHeight,
          baysWide: payload.baysWide,
          baysTall: payload.baysTall,
          customBayWidths: payload.customBayWidths,
          customBayHeights: payload.customBayHeights,
          doors: payload.doors,
          notes: payload.notes,
          createdAt: existing?.createdAt ?? now,
          updatedAt: now,
        }
        const updatedProject: Project = {
          ...project,
          elevations: existing
            ? project.elevations.map((el) => (el.id === elevationId ? elevation : el))
            : [elevation, ...project.elevations],
          updatedAt: now,
        }
        const projects = [...prev.projects]
        projects[projectIndex] = updatedProject
        result = {
          ok: true,
          message: `Elevation “${payload.name}” saved.`,
        }
        return {
          ...prev,
          projects,
          selectedProjectId: projectId,
          selectedElevationId: elevationId,
          status: makeStatus("success", result.message),
        }
      })
      return result
    },
    []
  )

  const deleteElevation = useCallback(
    (projectId: string, elevationId: string): ActionResult => {
      let result: ActionResult = { ok: false, message: "" }
      setState((prev) => {
        const projectIndex = prev.projects.findIndex(
          (proj) => proj.id === projectId
        )
        if (projectIndex === -1) {
          result = { ok: false, message: "Project not found." }
          return { ...prev, status: makeStatus("error", result.message) }
        }
        const project = prev.projects[projectIndex]
        const target = project.elevations.find((el) => el.id === elevationId)
        if (!target) {
          result = { ok: false, message: "Elevation not found." }
          return { ...prev, status: makeStatus("error", result.message) }
        }
        const updatedProject: Project = {
          ...project,
          elevations: project.elevations.filter((el) => el.id !== elevationId),
          updatedAt: new Date().toISOString(),
        }
        const projects = [...prev.projects]
        projects[projectIndex] = updatedProject
        const nextElevationId =
          prev.selectedElevationId === elevationId
            ? updatedProject.elevations[0]?.id
            : prev.selectedElevationId
        result = {
          ok: true,
          message: `Elevation “${target.name}” deleted.`,
        }
        return {
          ...prev,
          projects,
          selectedElevationId: nextElevationId,
          status: makeStatus("success", result.message),
        }
      })
      return result
    },
    []
  )

  const clearStatus = useCallback(() => {
    setState((prev) => ({ ...prev, status: undefined }))
  }, [])

  const selectedProject = useMemo(
    () => state.projects.find((proj) => proj.id === state.selectedProjectId),
    [state.projects, state.selectedProjectId]
  )

  const selectedElevation = useMemo(
    () => selectedProject?.elevations.find((el) => el.id === state.selectedElevationId),
    [selectedProject, state.selectedElevationId]
  )

  return {
    state,
    selectedProject,
    selectedElevation,
    actions: {
      createProject,
      deleteProject,
      selectProject,
      selectElevation,
      saveElevation,
      deleteElevation,
      clearStatus,
    },
  }
}
