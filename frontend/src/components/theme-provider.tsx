import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
} from "react"

type ThemeMode = "light" | "dark" | "system"
type ResolvedTheme = "light" | "dark"

interface ThemeContextValue {
  theme: ThemeMode
  resolvedTheme: ResolvedTheme
  setTheme: (theme: ThemeMode) => void
  toggleTheme: () => void
}

const ThemeContext = createContext<ThemeContextValue | null>(null)
const THEME_STORAGE_KEY = "estimation:theme"

const getSystemPreference = (): ResolvedTheme => {
  if (typeof window === "undefined" || !window.matchMedia) {
    return "light"
  }
  return window.matchMedia("(prefers-color-scheme: dark)").matches
    ? "dark"
    : "light"
}

const getInitialTheme = (): ThemeMode => {
  if (typeof window === "undefined") {
    return "light"
  }
  const stored = window.localStorage.getItem(THEME_STORAGE_KEY) as
    | ThemeMode
    | null
  return stored ?? "system"
}

export function ThemeProvider({ children }: { children: React.ReactNode }) {
  const [theme, setThemeState] = useState<ThemeMode>(() => getInitialTheme())
  const [systemTheme, setSystemTheme] = useState<ResolvedTheme>(() =>
    getSystemPreference()
  )

  const resolvedTheme: ResolvedTheme =
    theme === "system" ? systemTheme : theme

  useEffect(() => {
    if (typeof window === "undefined") return
    const mediaQuery = window.matchMedia("(prefers-color-scheme: dark)")
    const handler = () => setSystemTheme(mediaQuery.matches ? "dark" : "light")
    mediaQuery.addEventListener("change", handler)
    return () => mediaQuery.removeEventListener("change", handler)
  }, [])

  useEffect(() => {
    if (typeof document === "undefined") return
    document.documentElement.classList.toggle("dark", resolvedTheme === "dark")
    window.localStorage.setItem(THEME_STORAGE_KEY, theme)
  }, [resolvedTheme, theme])

  const setTheme = useCallback((next: ThemeMode) => {
    setThemeState(next)
  }, [])

  const toggleTheme = useCallback(() => {
    setThemeState((prev) => {
      const nextResolved =
        (prev === "system" ? systemTheme : prev) === "dark" ? "light" : "dark"
      if (prev === "system") {
        return nextResolved
      }
      return prev === "dark" ? "light" : "dark"
    })
  }, [systemTheme])

  const value = useMemo(
    () => ({
      theme,
      resolvedTheme,
      setTheme,
      toggleTheme,
    }),
    [theme, resolvedTheme, setTheme, toggleTheme]
  )

  return <ThemeContext.Provider value={value}>{children}</ThemeContext.Provider>
}

// eslint-disable-next-line react-refresh/only-export-components
export const useTheme = () => {
  const ctx = useContext(ThemeContext)
  if (!ctx) {
    throw new Error("useTheme must be used within ThemeProvider")
  }
  return ctx
}
