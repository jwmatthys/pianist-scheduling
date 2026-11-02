import { createContext, useCallback, useContext, useEffect, useState } from "react";
import type { ReactNode } from "react";
import { DEFAULT_THEME_PALETTE } from "./ThemePalette";
import { activateTheme, initializeTheme } from "./themeRuntime";
import type { ActiveTheme, ThemeSelection } from "./themeRuntime";

interface ThemeContextValue extends ActiveTheme {
  activate: (selection: ThemeSelection) => Promise<void>;
}

const ThemeContext = createContext<ThemeContextValue | null>(null);

export function ThemeProvider({ children }: { children: ReactNode }) {
  const [activeTheme, setActiveTheme] = useState<ActiveTheme>({
    selection: { id: "default" },
    palette: DEFAULT_THEME_PALETTE,
  });

  useEffect(() => {
    let active = true;
    void initializeTheme().then((restored) => {
      if (active) setActiveTheme({ selection: restored.selection, palette: restored.palette });
    }).catch((error: unknown) => {
      console.error("Theme restoration failed; retaining the built-in CSS theme.", error);
    });
    return () => {
      active = false;
    };
  }, []);

  const activate = useCallback(async (selection: ThemeSelection) => {
    const next = await activateTheme(selection);
    setActiveTheme(next);
  }, []);

  return (
    <ThemeContext.Provider value={{ ...activeTheme, activate }}>
      {children}
    </ThemeContext.Provider>
  );
}

export function useTheme(): ThemeContextValue {
  const context = useContext(ThemeContext);
  if (!context) throw new Error("useTheme must be used within a ThemeProvider.");
  return context;
}
