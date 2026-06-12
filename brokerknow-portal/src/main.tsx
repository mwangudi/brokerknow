import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { BrowserRouter } from "react-router";
import { HelmetProvider } from "react-helmet-async";
import { ThemeProvider } from "./context/ThemeContext";
import { AuthProvider } from "./context/AuthContext";
import App from "./App";
import "./index.css";

// Router base path. Mirrors Vite's `base` (import.meta.env.BASE_URL) so a
// sub-path deploy (e.g. `--base=/rw/`) and the default root deploy both work
// without further changes. Strip the trailing slash except for the root "/".
const basename = import.meta.env.BASE_URL.replace(/\/$/, "") || "/";

// Non-production builds (VITE_ENV_LABEL set) prefix the browser tab. Covers
// the login page, which has no per-page <title>; routed pages re-apply it via
// PageMeta. Empty in production = live tab title unchanged.
const envLabel = (import.meta.env.VITE_ENV_LABEL as string) || "";
if (envLabel) document.title = `[${envLabel}] ${document.title}`;

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <HelmetProvider>
      <ThemeProvider>
        <BrowserRouter basename={basename}>
          <AuthProvider>
            <App />
          </AuthProvider>
        </BrowserRouter>
      </ThemeProvider>
    </HelmetProvider>
  </StrictMode>,
);
