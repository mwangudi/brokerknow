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
    </>
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
