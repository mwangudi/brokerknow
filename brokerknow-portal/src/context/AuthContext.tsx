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
  /** Remaining "Not now" deferrals for a required change (server-capped at 3). */
  passwordChangeSkipsRemaining?: number;
}

interface AuthState {
  user: User | null;
  token: string | null;
  loading: boolean;
  login: (email: string, password: string) => Promise<{ error?: string; requiresPasswordChange?: boolean; otpRequired?: boolean; otpToken?: string; otpHint?: string }>;
  /** Completes a two-step login by verifying the emailed one-time passcode. */
  verifyOtp: (otpToken: string, code: string) => Promise<{ error?: string; requiresPasswordChange?: boolean }>;
  register: (data: RegisterData) => Promise<{ error?: string; message?: string }>;
  /** Resubmit a returned application via its emailed token. */
  resubmit: (token: string, data: RegisterData) => Promise<{ error?: string; message?: string }>;
  /** Request a password-reset code by email. Always returns a token for the code step. */
  forgotPassword: (email: string) => Promise<{ error?: string; message?: string; token?: string }>;
  /** Complete a password reset with the emailed code. */
  resetPassword: (token: string, code: string, newPassword: string) => Promise<{ error?: string; message?: string }>;
  logout: () => void;
  /** Marks the password change as completed and refreshes the cached user. */
  markPasswordChanged: () => void;
  /**
   * Defers a required password change ("Not now"). The server counts the
   * deferrals and refuses past the cap (3). Resolves with the remaining count;
   * rejects when no deferrals are left.
   */
  deferPasswordChange: () => Promise<number>;
}

interface RegisterData {
  email: string;
  firstName: string;
  lastName: string;
  phone?: string;
  officePhone?: string;
  homePhone?: string;
  idNumber?: string;
  idDocumentType?: string;
  cdsNumber?: string;
  dateOfBirth?: string;
  // Account-type intake. accountType = Individual|Joint|ITF. The joint/ITF
  // details travel as JSON strings the back office reads at approval time.
  accountType?: string;
  jointApplicants?: string;
  itfBeneficiary?: string;
  agreements?: string;
  physicalAddress?: string;
  postalAddress?: string;
  contactPerson?: string;
  nextOfKinName?: string;
  nextOfKinPhone?: string;
  nextOfKinAddress?: string;
  // True when an existing client is just requesting a portal login (skips KYC).
  isExistingClient?: boolean;
  // FIU / Cedar Capital KYC documents (required server-side for NEW applicants).
  // Each category may carry several files (e.g. corporates).
  idDocument?: File[];
  proofOfAddress?: File[];
  sourceOfFunds?: File[];
  otherDocuments?: File[];
  // A passport-style photo of the applicant (single image).
  passportPhoto?: File;
  // One ID document per joint account holder, in the same order as the
  // jointApplicants JSON so the server can map files to holders.
  jointIdDocuments?: File[];
}

// Build the multipart body shared by register + resubmit (KYC files travel
// with the scalar fields). Files are optional on resubmit (originals stay on
// file unless replaced).
function buildRegisterForm(data: RegisterData): FormData {
  const fd = new FormData();
  const scalarKeys = [
    "email", "firstName", "lastName", "phone", "officePhone", "homePhone",
    "idNumber", "idDocumentType", "cdsNumber", "dateOfBirth", "accountType",
    "jointApplicants", "itfBeneficiary", "agreements", "physicalAddress",
    "postalAddress", "contactPerson", "nextOfKinName", "nextOfKinPhone",
    "nextOfKinAddress",
  ] as const;
  for (const k of scalarKeys) {
    const v = data[k];
    if (v !== undefined && v !== null && v !== "") fd.append(k, String(v));
  }
  fd.append("isExistingClient", String(!!data.isExistingClient));
  if (data.idDocument) for (const f of data.idDocument) fd.append("idDocument", f, f.name);
  if (data.proofOfAddress) for (const f of data.proofOfAddress) fd.append("proofOfAddress", f, f.name);
  if (data.sourceOfFunds) for (const f of data.sourceOfFunds) fd.append("sourceOfFunds", f, f.name);
  if (data.otherDocuments) for (const f of data.otherDocuments) fd.append("otherDocuments", f, f.name);
  if (data.passportPhoto) fd.append("passportPhoto", data.passportPhoto, data.passportPhoto.name);
  if (data.jointIdDocuments) for (const f of data.jointIdDocuments) fd.append("jointIdDocuments", f, f.name);
  return fd;
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
      const r = await api.post("/auth/login", { email, password, audience: "portal" });
      if (r.data?.otpRequired) {
        return { otpRequired: true, otpToken: r.data.otpToken as string, otpHint: r.data.channelHint as string };
      }
      return applyAuth(r.data);
    } catch (err: any) {
      return { error: err.response?.data?.error || "Login failed." };
    } finally {
      setLoading(false);
    }
  }, []);

  const verifyOtp = useCallback(async (otpToken: string, code: string) => {
    setLoading(true);
    try {
      const r = await api.post("/auth/login/verify-otp", { otpToken, code });
      return applyAuth(r.data);
    } catch (err: any) {
      return { error: err.response?.data?.error || "Verification failed." };
    } finally {
      setLoading(false);
    }
  }, []);

  // Persist the issued tokens + user and flip the context into the signed-in
  // state. Shared by direct login and post-OTP verification.
  function applyAuth(data: any) {
    const { accessToken, refreshToken, user: u, requiresPasswordChange } = data;
    const userWithFlag: User = {
      ...u,
      mustChangePassword: !!requiresPasswordChange,
      passwordChangeSkipsRemaining: data.passwordChangeSkipsRemaining ?? 0,
    };
    localStorage.setItem("portal_token", accessToken);
    localStorage.setItem("portal_refresh", refreshToken);
    localStorage.setItem("portal_user", JSON.stringify(userWithFlag));
    setToken(accessToken);
    setUser(userWithFlag);
    return { requiresPasswordChange: !!requiresPasswordChange };
  }

  const register = useCallback(async (data: RegisterData) => {
    setLoading(true);
    try {
      // KYC documents must travel as multipart so the server can persist
      // them under /uploads/portal-applications/{userId}/.
      const fd = buildRegisterForm(data);
      // Pass undefined so axios fills in multipart/form-data with the
      // correct boundary; explicit \"multipart/form-data\" would drop it.
      const r = await api.post("/auth/register", fd, {
        headers: { "Content-Type": undefined as unknown as string },
      });
      return { message: r.data.message };
    } catch (err: any) {
      return { error: err.response?.data?.error || "Registration failed." };
    } finally {
      setLoading(false);
    }
  }, []);

  const resubmit = useCallback(async (token: string, data: RegisterData) => {
    setLoading(true);
    try {
      const fd = buildRegisterForm(data);
      const r = await api.post(`/auth/resubmit/${encodeURIComponent(token)}`, fd, {
        headers: { "Content-Type": undefined as unknown as string },
      });
      return { message: r.data.message };
    } catch (err: any) {
      return { error: err.response?.data?.error || "Resubmission failed." };
    } finally {
      setLoading(false);
    }
  }, []);

  const forgotPassword = useCallback(async (resetEmail: string) => {
    setLoading(true);
    try {
      const r = await api.post("/auth/forgot-password", { email: resetEmail });
      return { message: r.data.message as string, token: r.data.token as string };
    } catch (err: any) {
      return { error: err.response?.data?.error || "Could not start a password reset." };
    } finally {
      setLoading(false);
    }
  }, []);

  const resetPassword = useCallback(async (token: string, code: string, newPassword: string) => {
    setLoading(true);
    try {
      const r = await api.post("/auth/reset-password", { token, code, newPassword });
      return { message: r.data.message as string };
    } catch (err: any) {
      return { error: err.response?.data?.error || "Could not reset your password." };
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

  const deferPasswordChange = useCallback(async () => {
    const r = await api.post("/auth/defer-password-change");
    const remaining: number = r.data?.skipsRemaining ?? 0;
    setUser((prev) => {
      if (!prev) return prev;
      const next = { ...prev, mustChangePassword: false, passwordChangeSkipsRemaining: remaining };
      localStorage.setItem("portal_user", JSON.stringify(next));
      return next;
    });
    return remaining;
  }, []);

  return (
    <AuthContext.Provider value={{ user, token, loading, login, verifyOtp, register, resubmit, forgotPassword, resetPassword, logout, markPasswordChanged, deferPasswordChange }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error("useAuth must be used inside AuthProvider");
  return ctx;
}
