import axios from "axios";

const api = axios.create({
  baseURL: import.meta.env.VITE_API_URL || "http://localhost:5260/api",
  headers: { "Content-Type": "application/json" },
});

// Attach JWT token to every request
api.interceptors.request.use((config) => {
  const token = localStorage.getItem("portal_token");
  if (token) config.headers.Authorization = `Bearer ${token}`;
  return config;
});

/**
 * Absolute path to THIS deployment's login page, honouring the SPA basename
 * (import.meta.env.BASE_URL — e.g. "/", "/ke/", "/rw/"). Sub-path tenants must
 * bounce to their OWN login, not the Malawi instance at "/login".
 */
export function loginPath(): string {
  const base = import.meta.env.BASE_URL || "/"; // always ends with "/"
  return base.endsWith("/") ? `${base}login` : `${base}/login`;
}

// Auto-redirect on 401 (genuine session expiry only). We SKIP the auth
// endpoints: a 401 from /auth/login means "bad password" and a 401 from
// /auth/refresh means "refresh expired" — both should surface to the caller
// (inline error on the form) rather than bouncing the page to the login URL.
api.interceptors.response.use(
  (r) => r,
  (err) => {
    const url: string = err.config?.url ?? "";
    const isAuthEndpoint = url.includes("/auth/login") || url.includes("/auth/refresh");
    if (err.response?.status === 401 && !isAuthEndpoint) {
      localStorage.removeItem("portal_token");
      localStorage.removeItem("portal_refresh");
      localStorage.removeItem("portal_user");
      const target = loginPath();
      if (window.location.pathname !== target) {
        window.location.href = target;
      }
    }
    return Promise.reject(err);
  },
);

export default api;
