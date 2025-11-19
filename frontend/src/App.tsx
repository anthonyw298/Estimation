import { ThemeProvider } from "@/components/theme-provider"
import { EstimationDashboard } from "@/features/estimation/EstimationDashboard"

export default function App() {
  return (
    <ThemeProvider>
      <EstimationDashboard />
    </ThemeProvider>
  )
}
