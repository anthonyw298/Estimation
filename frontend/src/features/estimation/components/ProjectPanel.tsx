import { useMemo, useState } from "react"
import { Check, EllipsisVertical, Trash2 } from "lucide-react"

import { Button } from "@/components/ui/button"
import {
  Card,
  CardContent,
  CardDescription,
  CardFooter,
  CardHeader,
  CardTitle,
} from "@/components/ui/card"
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { ScrollArea } from "@/components/ui/scroll-area"
import { Badge } from "@/components/ui/badge"

import type { ActionResult, Project } from "../types"

interface ProjectPanelProps {
  projects: Project[]
  selectedProjectId?: string
  onCreateProject: (name: string) => ActionResult
  onSelectProject: (projectId: string) => void
  onDeleteProject: (projectId: string) => ActionResult
}

export function ProjectPanel({
  projects,
  selectedProjectId,
  onCreateProject,
  onSelectProject,
  onDeleteProject,
}: ProjectPanelProps) {
  const [projectName, setProjectName] = useState("")
  const [submitting, setSubmitting] = useState(false)

  const selectedProject = useMemo(
    () => projects.find((project) => project.id === selectedProjectId),
    [projects, selectedProjectId]
  )

  const handleCreateProject = (event: React.FormEvent) => {
    event.preventDefault()
    if (submitting) return
    setSubmitting(true)
    const result = onCreateProject(projectName)
    if (result.ok) {
      setProjectName("")
    }
    setSubmitting(false)
  }

  const handleDeleteProject = (projectId: string) => {
    onDeleteProject(projectId)
  }

  return (
    <div className="grid gap-6 lg:grid-cols-[360px,1fr]">
      <Card>
        <CardHeader>
          <CardTitle>Create project</CardTitle>
          <CardDescription>
            Capture individual jobs and organize their unique elevations.
          </CardDescription>
        </CardHeader>
        <CardContent>
          <form onSubmit={handleCreateProject} className="space-y-4">
            <div className="space-y-2">
              <Label htmlFor="project-name">Project name</Label>
              <Input
                id="project-name"
                placeholder="eg. Midtown Lobby Refresh"
                value={projectName}
                onChange={(event) => setProjectName(event.target.value)}
                required
              />
            </div>
            <Button type="submit" disabled={submitting}>
              {submitting ? "Adding…" : "Add project"}
            </Button>
          </form>
        </CardContent>
        <CardFooter>
          <p className="text-sm text-muted-foreground">
            Projects persist locally so you can return later without exporting.
          </p>
        </CardFooter>
      </Card>
      <Card className="flex flex-col">
        <CardHeader className="flex flex-row items-center justify-between gap-4">
          <div>
            <CardTitle>Project portfolio</CardTitle>
            <CardDescription>
              {projects.length
                ? "Select a project to manage its elevations."
                : "Start by creating your first project."}
            </CardDescription>
          </div>
          {selectedProject ? (
            <Badge variant="secondary" className="text-xs">
              Active: {selectedProject.name}
            </Badge>
          ) : null}
        </CardHeader>
        <CardContent className="flex-1">
          <ScrollArea className="h-[420px] pr-4">
            <div className="space-y-4">
              {projects.map((project) => (
                <div
                  key={project.id}
                  className="rounded-xl border bg-card/60 p-4 transition hover:border-primary/60"
                >
                  <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <div className="flex items-center gap-2 text-lg font-semibold">
                        {project.name}
                        {project.id === selectedProjectId ? (
                          <Badge variant="default" className="gap-1">
                            <Check className="h-3.5 w-3.5" />
                            Active
                          </Badge>
                        ) : null}
                      </div>
                      <p className="text-sm text-muted-foreground">
                        Updated {new Date(project.updatedAt).toLocaleString()}
                      </p>
                    </div>
                    <div className="flex items-center gap-2">
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => onSelectProject(project.id)}
                      >
                        Manage
                      </Button>
                      <DropdownMenu>
                        <DropdownMenuTrigger asChild>
                          <Button
                            variant="ghost"
                            size="icon"
                            aria-label="Project actions"
                          >
                            <EllipsisVertical className="h-4 w-4" />
                          </Button>
                        </DropdownMenuTrigger>
                        <DropdownMenuContent align="end">
                          <DropdownMenuItem
                            className="text-destructive"
                            onClick={() => handleDeleteProject(project.id)}
                          >
                            <Trash2 className="mr-2 h-4 w-4" />
                            Delete project
                          </DropdownMenuItem>
                        </DropdownMenuContent>
                      </DropdownMenu>
                    </div>
                  </div>
                  <div className="mt-3 flex flex-wrap items-center gap-3 text-sm">
                    <Badge variant="outline">
                      {project.elevations.length} elevation
                      {project.elevations.length === 1 ? "" : "s"}
                    </Badge>
                    {project.elevations[0] ? (
                      <span className="text-muted-foreground">
                        Latest elevation: {project.elevations[0].name}
                      </span>
                    ) : (
                      <span className="text-muted-foreground">
                        No elevations saved yet
                      </span>
                    )}
                  </div>
                </div>
              ))}
              {!projects.length ? (
                <div className="rounded-xl border border-dashed p-8 text-center text-muted-foreground">
                  No projects yet — your first one will appear here.
                </div>
              ) : null}
            </div>
          </ScrollArea>
        </CardContent>
      </Card>
    </div>
  )
}
