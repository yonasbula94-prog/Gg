import React, { useEffect, useState } from 'react';
import {
  CheckCircle2,
  CloudUpload,
  KeyRound,
  Laptop,
  RefreshCw,
  ShieldAlert,
  Trash2,
  UserPlus,
  Users,
} from 'lucide-react';
import {
  ROLE_LABELS,
  type OfflineLeaseToken,
  type OrgLicenseRecord,
  type OrgMemberRecord,
  type OrgRole,
  type OrganizationSummary,
} from '../types/enterprise';
import {
  activateOrgLicense,
  fetchOrgMembers,
  importLocalWorkspace,
  inviteOrgMember,
  removeOrgMember,
  updateOrgMember,
} from '../lib/apiClient';
import { loadCompanyProfile, loadJobsWithHealth, loadRegistryWithHealth } from '../store/storage';
import { isTauriDesktop } from '../utils/offlineSync';

interface OrgAdminViewProps {
  organization: OrganizationSummary;
  offlineLease: OfflineLeaseToken | null;
  pendingOutboxCount: number;
  onSyncNow: () => Promise<void>;
  onWorkspaceReloaded: () => Promise<void>;
  onNotify: (msg: string) => void;
}

export const OrgAdminView: React.FC<OrgAdminViewProps> = ({
  organization,
  offlineLease,
  pendingOutboxCount,
  onSyncNow,
  onWorkspaceReloaded,
  onNotify,
}) => {
  const [members, setMembers] = useState<OrgMemberRecord[]>([]);
  const [licenses, setLicenses] = useState<OrgLicenseRecord[]>([]);
  const [orgMeta, setOrgMeta] = useState(organization);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Invite form
  const [inviteEmail, setInviteEmail] = useState('');
  const [inviteName, setInviteName] = useState('');
  const [inviteRole, setInviteRole] = useState<OrgRole>('PROCUREMENT_OFFICER');
  const [inviting, setInviting] = useState(false);

  // License activation form
  const [licenseKeyInput, setLicenseKeyInput] = useState('');
  const [activatingKey, setActivatingKey] = useState(false);

  // Local workspace migration
  const [migrating, setMigrating] = useState(false);

  const isOrgAdmin = organization.role === 'ORG_ADMIN';
  const desktopMode = isTauriDesktop();

  const localStats = React.useMemo(() => {
    const j = loadJobsWithHealth();
    const r = loadRegistryWithHealth();
    return { jobCount: j.jobs.length, regCount: r.entries.length };
  }, []);

  const loadData = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetchOrgMembers(organization.id);
      setMembers(res.members);
      setLicenses(res.licenses);
      if (res.organization) {
        setOrgMeta({
          ...organization,
          ...res.organization,
          planTier: res.organization.planTier as OrganizationSummary['planTier'],
        });
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not load organization seats.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, [organization.id]);

  const seatsReserved = orgMeta.activeSeatsUsed + orgMeta.invitedSeatsUsed;
  const seatPct = Math.min(100, Math.round((seatsReserved / Math.max(1, orgMeta.seatLimit)) * 100));
  const atCapacity = seatsReserved >= orgMeta.seatLimit;

  const handleInvite = async (e: React.FormEvent) => {
    e.preventDefault();
    setInviting(true);
    setError(null);
    try {
      await inviteOrgMember(organization.id, {
        email: inviteEmail.trim(),
        displayName: inviteName.trim() || undefined,
        role: inviteRole,
      });
      setInviteEmail('');
      setInviteName('');
      onNotify(`Assigned ${ROLE_LABELS[inviteRole]} seat to ${inviteEmail.trim()}.`);
      await loadData();
      await onWorkspaceReloaded();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not assign seat.');
    } finally {
      setInviting(false);
    }
  };

  const handleActivateLicense = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!licenseKeyInput.trim()) return;
    setActivatingKey(true);
    setError(null);
    try {
      const res = await activateOrgLicense(organization.id, licenseKeyInput.trim());
      setLicenseKeyInput('');
      onNotify(
        `License activated: ${res.planTier} plan with ${res.seatLimit} seats valid until ${res.licenseExpiresAt.slice(0, 10)}.`
      );
      await loadData();
      await onWorkspaceReloaded();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not activate license key.');
    } finally {
      setActivatingKey(false);
    }
  };

  const handleRoleChange = async (member: OrgMemberRecord, nextRole: OrgRole) => {
    setError(null);
    try {
      await updateOrgMember(organization.id, member.id, { role: nextRole });
      onNotify(`Updated ${member.invitedEmail} role to ${ROLE_LABELS[nextRole]}.`);
      await loadData();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not update role.');
    }
  };

  const handleToggleSuspend = async (member: OrgMemberRecord) => {
    setError(null);
    const nextStatus = member.seatStatus === 'SUSPENDED' ? 'ACTIVE' : 'SUSPENDED';
    try {
      await updateOrgMember(organization.id, member.id, { seatStatus: nextStatus });
      onNotify(
        nextStatus === 'SUSPENDED'
          ? `Suspended seat for ${member.invitedEmail}. Active live sessions were disconnected.`
          : `Reactivated seat for ${member.invitedEmail}.`
      );
      await loadData();
      await onWorkspaceReloaded();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not change seat status.');
    }
  };

  const handleRemove = async (member: OrgMemberRecord) => {
    setError(null);
    try {
      await removeOrgMember(organization.id, member.id);
      onNotify(`Removed ${member.invitedEmail} and released 1 organization seat.`);
      await loadData();
      await onWorkspaceReloaded();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not remove member.');
    }
  };

  const handleMigrateLocalWorkspace = async () => {
    setMigrating(true);
    setError(null);
    try {
      const localJobs = loadJobsWithHealth().jobs;
      const localReg = loadRegistryWithHealth().entries;
      const localCompany = loadCompanyProfile();
      const res = await importLocalWorkspace(organization.id, {
        jobs: localJobs,
        registry: localReg,
        companyProfile: isOrgAdmin ? localCompany : undefined,
      });
      onNotify(
        `Migrated ${res.importedJobsCount} job(s) and ${res.importedRegistryCount} registry record(s) into ${organization.name}.`
      );
      await onWorkspaceReloaded();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not migrate local workspace.');
    } finally {
      setMigrating(false);
    }
  };

  return (
    <div className="space-y-6">
      {error && (
        <div className="rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-xs text-red-800 flex items-center gap-2">
          <ShieldAlert className="w-4 h-4 shrink-0" />
          <span>{error}</span>
        </div>
      )}

      {/* License & Seat Quota Overview */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        <section className="lg:col-span-7 bg-white border border-slate-200 rounded-xl p-6 space-y-5">
          <div className="flex flex-wrap items-start justify-between gap-4">
            <div>
              <h2 className="font-display text-lg font-bold text-slate-900">
                {orgMeta.name} — Seat License &amp; Quota
              </h2>
              <div className="text-xs text-slate-500 mt-1 flex flex-wrap items-center gap-2 font-mono">
                <span>Code: {orgMeta.orgCode}</span>
                <span aria-hidden="true">·</span>
                <span>Plan: {orgMeta.planTier}</span>
                <span aria-hidden="true">·</span>
                <span>
                  {orgMeta.licenseExpired
                    ? 'License Expired'
                    : `Expires ${orgMeta.licenseExpiresAt.slice(0, 10)} (${orgMeta.daysUntilExpiry}d left)`}
                </span>
              </div>
            </div>
            <button
              type="button"
              onClick={loadData}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700 cursor-pointer"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin' : ''}`} />
              <span>Refresh</span>
            </button>
          </div>

          <div className="space-y-2">
            <div className="flex items-center justify-between text-xs">
              <span className="font-medium text-slate-700">Assigned Organization Seats</span>
              <span className="font-mono tabular-nums font-semibold text-slate-900">
                {seatsReserved} / {orgMeta.seatLimit} seats used ({orgMeta.activeSeatsUsed} active ·{' '}
                {orgMeta.invitedSeatsUsed} invited)
              </span>
            </div>
            <div className="h-2.5 w-full rounded-full bg-slate-100 overflow-hidden">
              <div
                className={`h-full rounded-full transition-all ${
                  orgMeta.licenseExpired || atCapacity ? 'bg-amber-600' : 'bg-[#1B365D]'
                }`}
                style={{ width: `${seatPct}%` }}
              />
            </div>
          </div>

          {isOrgAdmin ? (
            <form onSubmit={handleActivateLicense} className="pt-4 border-t border-slate-200 space-y-3">
              <label
                htmlFor="upgrade-license-key"
                className="block text-xs font-semibold text-slate-800"
              >
                Apply or Upgrade Signed License Key
              </label>
              <div className="flex flex-col sm:flex-row gap-2.5">
                <input
                  id="upgrade-license-key"
                  type="text"
                  value={licenseKeyInput}
                  onChange={(e) => setLicenseKeyInput(e.target.value)}
                  placeholder="Paste signed GGPRO-... license key"
                  className="flex-1 min-h-[40px] rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none"
                />
                <button
                  type="submit"
                  disabled={activatingKey || !licenseKeyInput.trim()}
                  className="inline-flex items-center justify-center gap-1.5 px-4 py-2 rounded-lg bg-[#1B365D] hover:bg-[#152a48] text-white text-xs font-semibold cursor-pointer whitespace-nowrap disabled:opacity-50"
                >
                  <KeyRound className="w-3.5 h-3.5" />
                  <span>{activatingKey ? 'Verifying...' : 'Activate Key'}</span>
                </button>
              </div>
            </form>
          ) : (
            <div className="pt-3 border-t border-slate-200 text-xs text-slate-500">
              Only Organization Administrators can apply license keys or modify team seat allocations.
            </div>
          )}
        </section>

        {/* Desktop .exe & Offline Cache Status */}
        <section className="lg:col-span-5 bg-white border border-slate-200 rounded-xl p-6 space-y-4">
          <div className="flex items-start justify-between gap-3">
            <div>
              <h3 className="font-display text-base font-bold text-slate-900 flex items-center gap-2">
                <Laptop className="w-4 h-4 text-slate-500" />
                <span>Windows .exe &amp; Offline Sync</span>
              </h3>
              <p className="text-xs text-slate-500 mt-1">
                Runtime: <span className="font-semibold text-slate-800">{desktopMode ? 'Tauri Windows Desktop (.exe)' : 'Cloud Web Browser'}</span>
              </p>
            </div>
          </div>

          <div className="space-y-2.5 text-xs text-slate-600 border-t border-slate-100 pt-3">
            <div className="flex items-center justify-between">
              <span>Signed Offline Lease</span>
              <span className="font-mono tabular-nums text-slate-900">
                {offlineLease ? `Valid until ${offlineLease.expiresAt.slice(0, 10)}` : 'Active'}
              </span>
            </div>
            <div className="flex items-center justify-between">
              <span>Queued Offline Changes</span>
              <span className="font-mono tabular-nums font-semibold text-slate-900">
                {pendingOutboxCount}
              </span>
            </div>
            <div className="flex items-center justify-between">
              <span>Windows Installer Targets</span>
              <span className="font-mono text-slate-800">NSIS .exe · Enterprise .msi</span>
            </div>
          </div>

          <div className="pt-2 flex flex-wrap gap-2">
            <button
              type="button"
              onClick={onSyncNow}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700 cursor-pointer"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Sync Offline Outbox</span>
            </button>

            <button
              type="button"
              disabled={migrating}
              onClick={handleMigrateLocalWorkspace}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-medium text-slate-800 cursor-pointer disabled:opacity-60"
            >
              <CloudUpload className="w-3.5 h-3.5" />
              <span>
                {migrating
                  ? 'Migrating...'
                  : `Import Local Data (${localStats.jobCount} jobs, ${localStats.regCount} reg)`}
              </span>
            </button>
          </div>
        </section>
      </div>

      {/* Invite Seat Form (Org Admin Only) */}
      {isOrgAdmin && (
        <section className="bg-white border border-slate-200 rounded-xl p-6 space-y-4">
          <div className="flex items-center justify-between gap-4">
            <div>
              <h3 className="font-display text-base font-bold text-slate-900 flex items-center gap-2">
                <UserPlus className="w-4 h-4 text-slate-500" />
                <span>Assign New Organization Seat</span>
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                Assign the CEO role to stream the Executive Dashboard with full-app drill-down, or
                assign Procurement Officer seats for daily quotation and document workflows.
              </p>
            </div>
            {atCapacity && (
              <span className="text-xs font-semibold text-amber-700">
                Seat cap reached — apply an upgraded license key above
              </span>
            )}
          </div>

          <form onSubmit={handleInvite} className="grid grid-cols-1 md:grid-cols-12 gap-3 items-end">
            <div className="md:col-span-4">
              <label htmlFor="invite-email" className="block text-xs font-medium text-slate-700 mb-1">
                Work Email
              </label>
              <input
                id="invite-email"
                type="email"
                required
                value={inviteEmail}
                onChange={(e) => setInviteEmail(e.target.value)}
                placeholder="ceo@company.com"
                className="w-full min-h-[40px] rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none"
              />
            </div>

            <div className="md:col-span-3">
              <label htmlFor="invite-name" className="block text-xs font-medium text-slate-700 mb-1">
                Full Name
              </label>
              <input
                id="invite-name"
                type="text"
                value={inviteName}
                onChange={(e) => setInviteName(e.target.value)}
                placeholder="e.g. Chief Executive Officer"
                className="w-full min-h-[40px] rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none"
              />
            </div>

            <div className="md:col-span-3">
              <label htmlFor="invite-role" className="block text-xs font-medium text-slate-700 mb-1">
                Assigned Role
              </label>
              <select
                id="invite-role"
                value={inviteRole}
                onChange={(e) => setInviteRole(e.target.value as OrgRole)}
                className="w-full min-h-[40px] rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none"
              >
                <option value="CEO_EXECUTIVE">CEO / Executive (Live Dashboard + Full App)</option>
                <option value="PROCUREMENT_OFFICER">Procurement Officer</option>
                <option value="ORG_ADMIN">Organization Admin</option>
                <option value="VIEWER">Read-Only Viewer</option>
              </select>
            </div>

            <div className="md:col-span-2">
              <button
                type="submit"
                disabled={inviting || atCapacity || orgMeta.licenseExpired}
                className="w-full min-h-[40px] inline-flex items-center justify-center gap-1.5 rounded-lg bg-[#1B365D] hover:bg-[#152a48] text-white px-4 py-2 text-xs font-semibold cursor-pointer whitespace-nowrap disabled:opacity-50"
              >
                <span>{inviting ? 'Assigning...' : 'Assign Seat'}</span>
              </button>
            </div>
          </form>
        </section>
      )}

      {/* Team Members Seat Table */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between">
          <h3 className="font-display text-base font-bold text-slate-900 flex items-center gap-2">
            <Users className="w-4 h-4 text-slate-500" />
            <span>Organization Seats ({members.length})</span>
          </h3>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="border-b border-slate-200 bg-slate-50 text-[11px] font-semibold text-slate-500">
                <th className="py-2.5 px-4">Member</th>
                <th className="py-2.5 px-3">Role</th>
                <th className="py-2.5 px-3">Seat Status</th>
                <th className="py-2.5 px-3">Last Active</th>
                {isOrgAdmin && <th className="py-2.5 px-4 text-right">Seat Controls</th>}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-xs">
              {members.map((m) => (
                <tr key={m.id} className="hover:bg-slate-50/70">
                  <td className="py-3 px-4">
                    <div className="font-semibold text-slate-900">{m.displayName || m.invitedEmail}</div>
                    <div className="text-[11px] text-slate-500">{m.invitedEmail}</div>
                  </td>
                  <td className="py-3 px-3">
                    {isOrgAdmin ? (
                      <select
                        value={m.role}
                        onChange={(e) => handleRoleChange(m, e.target.value as OrgRole)}
                        className="rounded-md border border-slate-200 bg-white px-2 py-1 text-xs text-slate-800 focus:border-[#1B365D] focus:outline-none"
                      >
                        <option value="ORG_ADMIN">Organization Admin</option>
                        <option value="CEO_EXECUTIVE">CEO / Executive</option>
                        <option value="PROCUREMENT_OFFICER">Procurement Officer</option>
                        <option value="VIEWER">Read-Only Viewer</option>
                      </select>
                    ) : (
                      <span className="text-slate-700">{ROLE_LABELS[m.role]}</span>
                    )}
                  </td>
                  <td className="py-3 px-3 font-medium">
                    <span
                      className={
                        m.seatStatus === 'ACTIVE'
                          ? 'text-emerald-700'
                          : m.seatStatus === 'INVITED'
                          ? 'text-amber-700'
                          : 'text-red-700'
                      }
                    >
                      {m.seatStatus}
                    </span>
                  </td>
                  <td className="py-3 px-3 font-mono tabular-nums text-slate-500">
                    {m.lastActiveAt ? m.lastActiveAt.slice(0, 16).replace('T', ' ') : 'Pending login'}
                  </td>
                  {isOrgAdmin && (
                    <td className="py-3 px-4 text-right whitespace-nowrap space-x-2">
                      <button
                        type="button"
                        onClick={() => handleToggleSuspend(m)}
                        className="px-2.5 py-1 rounded-md border border-slate-200 bg-white hover:bg-slate-100 text-[11px] font-medium text-slate-700 cursor-pointer"
                      >
                        {m.seatStatus === 'SUSPENDED' ? 'Reactivate' : 'Suspend'}
                      </button>
                      <button
                        type="button"
                        onClick={() => handleRemove(m)}
                        title="Remove seat"
                        className="p-1.5 rounded-md text-slate-400 hover:text-red-600 hover:bg-red-50 cursor-pointer align-middle"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {/* Activated License History */}
      {licenses.length > 0 && (
        <section className="bg-white border border-slate-200 rounded-xl p-5 space-y-3">
          <h3 className="font-display text-sm font-bold text-slate-900 flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 text-emerald-600" />
            <span>Activated License History</span>
          </h3>
          <div className="divide-y divide-slate-100 text-xs">
            {licenses.map((lic) => (
              <div key={lic.id} className="py-2.5 flex flex-wrap items-center justify-between gap-3">
                <div className="space-y-0.5">
                  <div className="font-mono font-semibold text-slate-900">
                    {lic.id} · {lic.planTier} ({lic.seatLimit} seats)
                  </div>
                  <div className="text-[11px] text-slate-500">
                    Valid until {lic.expiresAt.slice(0, 10)}
                    {lic.activatedAt ? ` · Activated ${lic.activatedAt.slice(0, 10)}` : ''}
                  </div>
                </div>
                <span className="font-mono text-[11px] text-emerald-700 font-semibold">
                  {lic.status}
                </span>
              </div>
            ))}
          </div>
        </section>
      )}
    </div>
  );
};
