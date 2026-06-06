import { useEffect, useState } from "react";
import api from "../lib/api";
import { useAuth } from "../context/AuthContext";
import PageMeta from "../components/common/PageMeta";

interface Profile {
  clientDpa: number;
  clientName: string;
  clientIdPass?: string;
  clientCdsNo?: string;
  clientEmail?: string;
  clientCellTel?: string;
  clientOfficeTel?: string;
  clientHomeTel?: string;
  clientAddr?: string;
  clientPAddr?: string;
  clientContact?: string;
  clientBDate?: string;
}

export default function ProfilePage() {
  const { user } = useAuth();
  const [profile, setProfile] = useState<Profile | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    api
      .get("/portal/profile")
      .then((r) => setProfile(r.data))
      .catch((err) => setError(err.response?.data?.error || "Failed to load profile."))
      .finally(() => setLoading(false));
  }, []);

  if (loading) {
    return (
      <div className="flex h-64 items-center justify-center text-gray-500 dark:text-gray-400">
        Loading profile...
      </div>
    );
  }
  if (error) {
    return (
      <div className="rounded-2xl border border-error-200 bg-error-50 p-6 text-error-700 dark:border-error-500/30 dark:bg-error-500/10 dark:text-error-400">
        {error}
      </div>
    );
  }
  if (!profile) return null;

  const initials = `${profile.clientName?.[0] ?? "C"}`.toUpperCase();

  return (
    <>
      <PageMeta title="My Profile" description="View and manage your profile" />

      {/* Hero card */}
      <div className="mb-6 rounded-2xl border border-gray-200 bg-white p-6 dark:border-gray-800 dark:bg-white/[0.03]">
        <div className="flex items-center gap-5">
          <div className="flex h-20 w-20 shrink-0 items-center justify-center rounded-full bg-brand-50 text-2xl font-bold text-brand-600 dark:bg-brand-500/10 dark:text-brand-400">
            {initials}
          </div>
          <div className="min-w-0 flex-1">
            <h2 className="truncate text-2xl font-bold text-gray-900 dark:text-white">
              {profile.clientName}
            </h2>
            <p className="mt-1 text-sm text-gray-500 dark:text-gray-400">
              Client ID: <span className="font-medium text-gray-700 dark:text-gray-300">#{profile.clientDpa}</span>
            </p>
            <p className="text-xs text-gray-400 mt-0.5">
              Logged in as {user?.email}
            </p>
          </div>
        </div>
      </div>

      {/* Detail cards */}
      <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
        <DetailCard
          title="Personal Information"
          icon={
            <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
            </svg>
          }
          gradient="bg-blue-50 text-blue-600 dark:bg-blue-500/10 dark:text-blue-400"
          rows={[
            { label: "ID / Passport", value: profile.clientIdPass },
            { label: "CDS Number", value: profile.clientCdsNo },
            { label: "Date of Birth", value: profile.clientBDate ? new Date(profile.clientBDate).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : undefined },
            { label: "Contact Person", value: profile.clientContact },
          ]}
        />
        <DetailCard
          title="Contact Details"
          icon={
            <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z" />
            </svg>
          }
          gradient="bg-success-50 text-success-600 dark:bg-success-500/10 dark:text-success-400"
          rows={[
            { label: "Email", value: profile.clientEmail },
            { label: "Cell Phone", value: profile.clientCellTel },
            { label: "Office Phone", value: profile.clientOfficeTel },
            { label: "Home Phone", value: profile.clientHomeTel },
          ]}
        />
        <DetailCard
          title="Addresses"
          icon={
            <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17.657 16.657L13.414 20.9a2 2 0 01-2.828 0l-4.244-4.243a8 8 0 1111.314 0z" />
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 11a3 3 0 11-6 0 3 3 0 016 0z" />
            </svg>
          }
          gradient="bg-warning-50 text-warning-600 dark:bg-warning-500/10 dark:text-warning-400"
          className="lg:col-span-2"
          rows={[
            { label: "Physical Address", value: profile.clientAddr },
            { label: "Postal Address", value: profile.clientPAddr },
          ]}
        />
      </div>

      {/* Documents — KYC / account-opening files uploaded by the back-office.
          Read-only: portal clients can download but not upload or delete. */}
      <div className="mt-6">
        <DocumentsCard />
      </div>
    </>
  );
}

interface PortalAttachment {
  id: string;
  name: string;
  size: number;
  uploadedAt: string;
  contentType: string;
}

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(2)} MB`;
}

function DocumentsCard() {
  const [items, setItems] = useState<PortalAttachment[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    api
      .get<PortalAttachment[]>("/portal/attachments")
      .then((r) => {
        if (!cancelled) setItems(r.data ?? []);
      })
      .catch((err) => {
        if (!cancelled) {
          setError(err.response?.data?.error || "Failed to load documents.");
        }
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  async function download(a: PortalAttachment) {
    try {
      const res = await api.get(`/portal/attachments/${encodeURIComponent(a.id)}`, {
        responseType: "blob",
      });
      const url = window.URL.createObjectURL(res.data as Blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = a.name;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      const e = err as { message?: string };
      setError(e.message ?? "Download failed.");
    }
  }

  return (
    <div className="rounded-2xl border border-gray-200 bg-white dark:border-gray-800 dark:bg-white/[0.03]">
      <div className="flex items-center gap-3 border-b border-gray-100 px-6 py-4 dark:border-gray-800">
        <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-purple-50 text-purple-600 dark:bg-purple-500/10 dark:text-purple-400">
          <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2M9 12h6M9 16h6" />
          </svg>
        </div>
        <div className="min-w-0 flex-1">
          <h3 className="font-semibold text-gray-900 dark:text-white">My Documents</h3>
          <p className="text-xs text-gray-500 dark:text-gray-400">
            KYC and account-opening files. Contact your account manager if anything is missing.
          </p>
        </div>
      </div>

      <div className="p-6">
        {error && (
          <div className="mb-4 rounded-lg border border-error-200 bg-error-50 px-3 py-2 text-xs text-error-700 dark:border-error-500/30 dark:bg-error-500/10 dark:text-error-400">
            {error}
          </div>
        )}

        {loading ? (
          <p className="text-sm text-gray-500 dark:text-gray-400">Loading documents...</p>
        ) : items.length === 0 ? (
          <div className="rounded-lg border border-dashed border-gray-300 bg-gray-50 px-4 py-8 text-center text-sm text-gray-500 dark:border-gray-700 dark:bg-white/[0.02] dark:text-gray-400">
            No documents on file yet.
          </div>
        ) : (
          <ul className="divide-y divide-gray-100 dark:divide-gray-800">
            {items.map((a) => (
              <li key={a.id} className="flex items-center gap-3 py-3">
                <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-brand-50 text-brand-500 dark:bg-brand-500/15 dark:text-brand-400">
                  <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" />
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M14 2v6h6" />
                  </svg>
                </div>
                <div className="min-w-0 flex-1">
                  <button
                    type="button"
                    onClick={() => download(a)}
                    className="block truncate text-left text-sm font-medium text-gray-800 hover:text-brand-600 hover:underline dark:text-white/90 dark:hover:text-brand-400"
                    title={a.name}
                  >
                    {a.name}
                  </button>
                  <div className="text-xs text-gray-500 dark:text-gray-400">
                    {formatBytes(a.size)} · uploaded{" "}
                    {new Date(a.uploadedAt).toLocaleDateString("en-GB", {
                      day: "2-digit",
                      month: "short",
                      year: "numeric",
                    })}
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => download(a)}
                  className="rounded border border-gray-300 bg-white px-2.5 py-1 text-xs text-gray-700 hover:bg-gray-50 dark:border-gray-700 dark:bg-gray-800 dark:text-gray-300"
                >
                  Download
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}

function DetailCard({
  title, icon, gradient, rows, className = "",
}: {
  title: string;
  icon: React.ReactNode;
  gradient: string;
  rows: { label: string; value?: string | null }[];
  className?: string;
}) {
  return (
    <div className={`rounded-2xl border border-gray-200 bg-white dark:border-gray-800 dark:bg-white/[0.03] ${className}`}>
      <div className="flex items-center gap-3 border-b border-gray-100 px-6 py-4 dark:border-gray-800">
        <div className={`flex h-9 w-9 items-center justify-center rounded-lg ${gradient}`}>
          {icon}
        </div>
        <h3 className="font-semibold text-gray-900 dark:text-white">{title}</h3>
      </div>
      <dl className="grid grid-cols-1 gap-x-6 gap-y-4 p-6 sm:grid-cols-2">
        {rows.map((r) => (
          <div key={r.label}>
            <dt className="text-xs font-medium uppercase tracking-wider text-gray-500">{r.label}</dt>
            <dd className="mt-1 text-sm text-gray-900 dark:text-white/90">
              {r.value || <span className="text-gray-400 italic">Not provided</span>}
            </dd>
          </div>
        ))}
      </dl>
    </div>
  );
}
