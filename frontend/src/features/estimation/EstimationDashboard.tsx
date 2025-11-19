import { useState } from "react"
import { Building2, ClipboardList, LayoutPanelTop } from "lucide-react"

import { ThemeToggle } from "@/components/theme-toggle"
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert"
import { Badge } from "@/components/ui/badge"
import { Button } from "@/components/ui/button"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"

import { ProjectPanel } from "./components/ProjectPanel"
import { ElevationWorkspace } from "./components/ElevationWorkspace"
import { useEstimationState } from "./useEstimationState"

export function EstimationDashboard() {
  const { state, selectedProject, selectedElevation, actions } = useEstimationState()
  const [activeTab, setActiveTab] = useState<"projects" | "elevations">("projects")

  return (
    <div className="min-h-screen bg-background">
      <div className="mx-auto flex max-w-7xl flex-col gap-6 px-6 py-8">
        <header className="flex flex-col gap-6 rounded-3xl border bg-card/40 p-6 shadow-sm shadow-primary/5 md:flex-row md:items-center md:justify-between">
          <div className="space-y-2">
            <div className="flex items-center gap-2 text-sm text-muted-foreground">
              <Building2 className="h-4 w-4 text-primary" />
              Estimation command center
            </div>
            <h1 className="text-3xl font-semibold tracking-tight">
              United Glass Estimator
            </h1>
            <p className="max-w-3xl text-sm text-muted-foreground">
              Manage storefront projects, curate elevations, and iterate on door & hardware packages with a fluid, modern UI.
            </p>
            <div className="flex flex-wrap gap-2">
              <Badge variant="outline">{state.projects.length} project(s)</Badge>
              <Badge variant="secondary">
                {selectedProject ? selectedProject.name : "No project selected"}
              </Badge>
            </div>
          </div>
          <div className="flex items-center gap-3 self-start md:self-auto">
            <Button variant="outline" className="gap-2" onClick={() => setActiveTab("projects")}>
              <ClipboardList className="h-4 w-4" />
              Projects
            </Button>
            <Button
              variant="outline"
              className="gap-2"
              onClick={() => setActiveTab("elevations")}
              disabled={!selectedProject}
            >
              <LayoutPanelTop className="h-4 w-4" />
              Elevations
            </Button>
            <ThemeToggle />
          </div>
        </header>

        {state.status ? (
          <Alert variant={state.status.tone === "error" ? "destructive" : "default"}>
            <AlertTitle className="capitalize">{state.status.tone}</AlertTitle>
            <AlertDescription className="flex items-center justify-between gap-4">
              <span>{state.status.text}</span>
              <Button
                variant="ghost"
                size="sm"
                onClick={actions.clearStatus}
                className="text-muted-foreground hover:text-foreground"
              >
                Dismiss
              </Button>
            </AlertDescription>
          </Alert>
        ) : null}

        <Tabs
          value={activeTab}
          onValueChange={(value) => setActiveTab(value as typeof activeTab)}
          className="space-y-6"
        >
          <TabsList className="w-full justify-start">
            <TabsTrigger value="projects">Projects</TabsTrigger>
            <TabsTrigger value="elevations" disabled={!selectedProject}>
              Elevations
            </TabsTrigger>
          </TabsList>
          <TabsContent value="projects" className="space-y-6">
            <ProjectPanel
              projects={state.projects}
              selectedProjectId={state.selectedProjectId}
              onCreateProject={actions.createProject}
              onSelectProject={(projectId) => {
                actions.selectProject(projectId)
                setActiveTab("elevations")
              }}
              onDeleteProject={actions.deleteProject}
            />
          </TabsContent>
          <TabsContent value="elevations">
            <ElevationWorkspace
              project={selectedProject}
              selectedElevation={selectedElevation}
              onSelectElevation={actions.selectElevation}
              onSaveElevation={actions.saveElevation}
              onDeleteElevation={actions.deleteElevation}
            />
          </TabsContent>
        </Tabs>
      </div>
    </div>
  )
}
