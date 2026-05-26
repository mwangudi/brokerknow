import { createContext, useContext, useState, useCallback, useEffect, type ReactNode } from "react";
import api from "../lib/api";

interface User {
  id: number;
  email: string;
  firstName: string;
  lastName: string;
  phone?: string;
  idNumber?: string;
  role: string;
  status: string;
  clientDpa?: number;
  /** True when the current password is temporary or has expired. */
  mustChangePassword?: boolean;
  passwordChangedAt?: string | null;
}

interface AuthState {
  user: User | null;
  token: string | null;
  loading: boolean;
  login: (email: string, password: string) => Promise<{ error?: string; requiresPasswordChange?: boolean }>;
  register: (data: RegisterData) => Promise<{ error?: string; message?: string }>;
  logout: () => void;
  /** Marks the password change as completed and refreshes the cached user. */
  markPasswordChanged: () => void;
}

interface RegisterData {
  email: string;
  firstName: string;
  lastName: string;
  phone?: string;
  officePhone?: string;
  homePhone?: string;
  idNumber?: string;
  cdsNumber?: string;
  dateOfBirth?: string;
  physicalAddress?: string;
  postalAddress?: string;
  contactPerson?: string;
}

const AuthContext = createContext<AuthState | null>(null);

export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<User | null>(() => {
    const stored = localStorage.getItem("portal_user");
    return stored ? JSON.parse(stored) : null;
  });
  const [token, setToken] = useState<string | null>(
    () => localStorage.getItem("portal_token"),
  );
  const [loading, setLoading] = useState(false);

  // Verify token on mount
  useEffect(() => {
    if (token && !user) {
      api.get("/auth/me")
        .then((r) => {
          setUser(r.data);
          localStorage.setItem("portal_user", JSON.stringify(r.data));
        })
        .catch(() => {
          setToken(null);
          localStorage.removeItem("portal_token");
          localStorage.removeItem("portal_user");
        });
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const login = useCallback(async (email: string, password: string) => {
    setLoading(true);
    try {
      const r = await api.post("/auth/login", { email, password });
      const { accessToken, refreshToken, user: u, requiresPasswordChange } = r.data;
      // Reflect the server's policy decision on the cached user object.
      const userWithFlag: User = { ...u, mustChangePassword: !!requiresPasswordChange };
      localStorage.setItem("portal_token", accessToken);
      localStorage.setItem("portal_refresh", refreshToken);
      localStorage.setItem("portal_user", JSON.stringify(userWithFlag));
      setToken(accessToken);
      setUser(userWithFlag);
      return { requiresPasswordChange: !!requiresPasswordChange };
    } catch (err: any) {
      return { error: err.response?.data?.error || "Login failed." };
    } finally {
      setLoading(false);
    }
  }, []);

  const register = useCallback(async (data: RegisterData) => {
    setLoading(true);
    try {
      const r = await api.post("/auth/register", data);
      return { message: r.data.message };
    } catch (err: any) {
      return { error: err.response?.data?.error || "Registration failed." };
    } finally {
      setLoading(false);
    }
  }, []);

  const logout = useCallback(() => {
    const refresh = localStorage.getItem("portal_refresh");
    if (refresh) api.post("/auth/logout", { refreshToken: refresh }).catch(() => {});
    localStorage.removeItem("portal_token");
    localStorage.removeItem("portal_refresh");
    localStorage.removeItem("portal_user");
    setToken(null);
    setUser(null);
  }, []);

  const markPasswordChanged = useCallback(() => {
    setUser((prev) => {
      if (!prev) return prev;
      const next = { ...prev, mustChangePassword: false, passwordChangedAt: new Date().toISOString() };
      localStorage.setItem("portal_user", JSON.stringify(next));
      return next;
    });
  }, []);

  return (
    <AuthContext.Provider value={{ user, token, loading, login, register, logout, markPasswordChanged }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error("useAuth must be used inside AuthProvider");
  return ctx;
}
