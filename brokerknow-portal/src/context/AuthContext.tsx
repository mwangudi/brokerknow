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
  login: (email: string, password: string) => Promise<{ error?: string; requiresPasswordChange?: boolean; otpRequired?: boolean; otpToken?: string; otpHint?: string }>;
  /** Completes a two-step login by verifying the emailed one-time passcode. */
  verifyOtp: (otpToken: string, code: string) => Promise<{ error?: string; requiresPasswordChange?: boolean }>;
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
  idDocumentType?: string;
  cdsNumber?: string;
  dateOfBirth?: string;
  // Account-type intake. accountType = Individual|Joint|ITF. The joint/ITF
  // details travel as JSON strings the back office reads at approval time.
  accountType?: string;
  jointApplicants?: string;
  itfBeneficiary?: string;
  physicalAddress?: string;
  postalAddress?: string;
  contactPerson?: string;
  // True when an existing client is just requesting a portal login (skips KYC).
  isExistingClient?: boolean;
  // FIU / Cedar Capital KYC documents (required server-side for NEW applicants).
  idDocument?: File;
  proofOfAddress?: File;
  sourceOfFunds?: File;
  otherDocuments?: File[];
  // One ID document per joint account holder, in the same order as the
  // jointApplicants JSON so the server can map files to holders.
  jointIdDocuments?: File[];
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
    const userWithFlag: User = { ...u, mustChangePassword: !!requiresPasswordChange };
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
      const fd = new FormData();
      const scalarKeys = [
        "email", "firstName", "lastName", "phone", "officePhone", "homePhone",
        "idNumber", "idDocumentType", "cdsNumber", "dateOfBirth", "accountType",
        "jointApplicants", "itfBeneficiary", "physicalAddress",
        "postalAddress", "contactPerson",
      ] as const;
      for (const k of scalarKeys) {
        const v = data[k];
        if (v !== undefined && v !== null && v !== "") fd.append(k, String(v));
      }
      // Always send the existing-client flag so the server can relax the KYC
      // upload + duplicate-client checks for returning clients.
      fd.append("isExistingClient", String(!!data.isExistingClient));
      if (data.idDocument) fd.append("idDocument", data.idDocument, data.idDocument.name);
      if (data.proofOfAddress) fd.append("proofOfAddress", data.proofOfAddress, data.proofOfAddress.name);
      if (data.sourceOfFunds) fd.append("sourceOfFunds", data.sourceOfFunds, data.sourceOfFunds.name);
      if (data.otherDocuments) {
        for (const f of data.otherDocuments) fd.append("otherDocuments", f, f.name);
      }
      if (data.jointIdDocuments) {
        for (const f of data.jointIdDocuments) fd.append("jointIdDocuments", f, f.name);
      }
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
    <AuthContext.Provider value={{ user, token, loading, login, verifyOtp, register, logout, markPasswordChanged }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error("useAuth must be used inside AuthProvider");
  return ctx;
}
