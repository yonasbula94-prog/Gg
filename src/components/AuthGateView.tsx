import React, { useState } from 'react';
import {
  ArrowRight,
  Building2,
  CheckCircle2,
  KeyRound,
  LogOut,
  Monitor,
  ShieldCheck,
  Users,
  WifiOff,
} from 'lucide-react';
import { ROLE_LABELS, type AuthUserProfile, type OrganizationSummary, type PendingInvite } from '../types/enterprise';
import { acceptInvite, createOrganization } from '../lib/apiClient';
import { isTauriDesktop } from '../utils/offlineSync';

interface AuthGateViewProps {
  mode: 'sign_in' | 'org_onboarding';
  user: AuthUserProfile | null;
  pendingInvites: PendingInvite[];
  authLoading: boolean;
  authError: string | null;
  onGoogleSignIn: () => Promise<void>;
  onSignOut: () => Promise<void>;
  onEnterLocalOfflineMode: () => void;
  onOrganizationReady: (org: OrganizationSummary) => Promise<void>;
  onOpenVendorConsole?: () => void;
}

export const AuthGateView: React.FC<AuthGateViewProps> = ({
  mode,
  user,
  pendingInvites,
  authLoading,
  authError,
  onGoogleSignIn,
  onSignOut,
  onEnterLocalOfflineMode,
  onOrganizationReady,
  onOpenVendorConsole,
}) => {
  const [orgName, setOrgName] = useState('');
  const [licenseKey, setLicenseKey] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [acceptingId, setAcceptingId] = useState<string | null>(null);
  const [formError, setFormError] = useState<string | null>(null);

  const desktopRuntime = isTauriDesktop();

  const handleCreateOrg = async (e: React.FormEvent) => {
    e.preventDefault();
    setFormError(null);
    if (orgName.trim().length < 2) {
      setFormError('Enter your organization name (at least 2 characters).');
      return;
    }
    setSubmitting(true);
    try {
      const res = await createOrganization({
        name: orgName.trim(),
        licenseKey: licenseKey.trim() || undefined,
      });
      if (!res?.organization?.id) {
        throw new Error('Organization workspace response was incomplete. Please try again.');
      }
      await onOrganizationReady(res.organization);
    } catch (err) {
      setFormError(err instanceof Error ? err.message : 'Could not create organization workspace.');
    } finally {
      setSubmitting(false);
    }
  };

  const handleAcceptInvite = async (invite: PendingInvite) => {
    setFormError(null);
    setAcceptingId(invite.membershipId);
    try {
      const res = await acceptInvite(invite.membershipId);
      if (!res?.organization?.id) {
        throw new Error('Invitation response was incomplete. Please try again.');
      }
      await onOrganizationReady(res.organization);
    } catch (err) {
      setFormError(err instanceof Error ? err.message : 'Could not accept invitation.');
    } finally {
      setAcceptingId(null);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col">
      {/* Top Bar */}
      <header className="h-16 border-b border-slate-200 bg-white px-6 flex items-center justify-between gap-8">
        <div className="flex items-center gap-3 min-w-0">
          <div className="w-8 h-8 rounded-lg bg-[#1B365D] text-white flex items-center justify-center font-display font-bold text-sm shrink-0">
            GG
          </div>
          <span className="font-display text-base font-bold tracking-tight text-slate-900 whitespace-nowrap">
            Grand Group Procurement Suite
          </span>
        </div>

        <div className="flex items-center gap-4 text-xs text-slate-500">
          <span className="hidden sm:inline whitespace-nowrap">
            {desktopRuntime ? 'Windows Desktop Client' : 'Cloud Multi-User Workspace'}
          </span>
          {user && (
            <button
              type="button"
              onClick={onSignOut}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 font-medium cursor-pointer whitespace-nowrap"
            >
              <LogOut className="w-3.5 h-3.5" />
              <span>Sign Out</span>
            </button>
          )}
        </div>
      </header>

      {/* Main Split Content */}
      <main className="flex-1 max-w-6xl w-full mx-auto px-6 py-12 grid grid-cols-1 lg:grid-cols-12 gap-10 items-stretch">
        {/* Left: Value & Architecture Overview */}
        <section className="lg:col-span-5 flex flex-col justify-between bg-[#1B365D] text-white rounded-xl p-8 lg:p-10">
          <div className="space-y-6">
            <div className="text-xs font-medium text-slate-300 tracking-wide">
              Enterprise Procurement &amp; Trade Operations
            </div>
            <h1 className="font-display text-3xl font-bold leading-tight text-balance">
              Multi-user procurement workspaces with live executive visibility.
            </h1>
            <p className="text-sm text-slate-200 leading-relaxed">
              Coordinate supplier RFQs, landed-cost pricing, all six trade documents, and an atomic
              Master Registry across your Windows desktop and cloud team seats.
            </p>

            <div className="pt-4 space-y-4 border-t border-white/15">
              <div className="flex items-start gap-3">
                <Users className="w-4 h-4 text-amber-300 shrink-0 mt-0.5" />
                <div>
                  <div className="text-sm font-semibold">Seat-Licensed Organization Access</div>
                  <div className="text-xs text-slate-300 mt-0.5">
                    Role-based seats for Admins, CEOs, Procurement Officers, and Viewers backed by
                    signed license keys.
                  </div>
                </div>
              </div>

              <div className="flex items-start gap-3">
                <ShieldCheck className="w-4 h-4 text-amber-300 shrink-0 mt-0.5" />
                <div>
                  <div className="text-sm font-semibold">Collision-Free Master Registry</div>
                  <div className="text-xs text-slate-300 mt-0.5">
                    PostgreSQL transactional numbering guarantees sequential QT, PI, PO, CI, PL, and
                    DN codes across every workstation.
                  </div>
                </div>
              </div>

              <div className="flex items-start gap-3">
                <Monitor className="w-4 h-4 text-amber-300 shrink-0 mt-0.5" />
                <div>
                  <div className="text-sm font-semibold">Live CEO Dashboard &amp; Windows .exe</div>
                  <div className="text-xs text-slate-300 mt-0.5">
                    Real-time pipeline streaming for executives plus Tauri offline caching for field
                    procurement teams.
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div className="pt-8 mt-8 border-t border-white/15 text-xs text-slate-300 flex items-center justify-between">
            <span>PostgreSQL Cloud Sync</span>
            <span className="font-mono tabular-nums">v2.0 Enterprise</span>
          </div>
        </section>

        {/* Right: Sign-In or Organization Activation */}
        <section className="lg:col-span-7 bg-white border border-slate-200 rounded-xl p-8 lg:p-10 flex flex-col justify-center">
          {mode === 'sign_in' ? (
            <div className="max-w-md mx-auto w-full space-y-6">
              <div>
                <h2 className="font-display text-2xl font-bold text-slate-900">
                  Sign in to your workspace
                </h2>
                <p className="text-sm text-slate-600 mt-1.5">
                  Use your organization Google account to access your assigned seat, shared Master
                  Registry, and live executive stream.
                </p>
              </div>

              {authError && (
                <div className="rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-xs text-red-800">
                  {authError}
                </div>
              )}

              <button
                type="button"
                disabled={authLoading}
                onClick={onGoogleSignIn}
                className="w-full min-h-[44px] flex items-center justify-center gap-3 rounded-lg bg-[#1B365D] hover:bg-[#152a48] text-white px-5 py-3 text-sm font-semibold transition-colors cursor-pointer disabled:opacity-60"
              >
                <span>{authLoading ? 'Signing in...' : 'Continue with Google'}</span>
                <ArrowRight className="w-4 h-4" />
              </button>

              <div className="relative py-2">
                <div className="absolute inset-0 flex items-center">
                  <div className="w-full border-t border-slate-200" />
                </div>
                <div className="relative flex justify-center text-xs">
                  <span className="bg-white px-3 text-slate-400">or work locally</span>
                </div>
              </div>

              <button
                type="button"
                onClick={onEnterLocalOfflineMode}
                className="w-full min-h-[42px] flex items-center justify-center gap-2 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700 px-4 py-2.5 text-xs font-medium transition-colors cursor-pointer"
              >
                <WifiOff className="w-4 h-4 text-slate-500" />
                <span>Open Standalone Offline Local Mode (Single Device)</span>
              </button>
            </div>
          ) : (
            <div className="space-y-8">
              <div className="flex items-start justify-between gap-4 border-b border-slate-200 pb-5">
                <div>
                  <h2 className="font-display text-2xl font-bold text-slate-900">
                    Connect to an organization
                  </h2>
                  <p className="text-sm text-slate-600 mt-1">
                    Signed in as <span className="font-medium text-slate-900">{user?.email}</span>
                  </p>
                </div>
                {user?.isVendorAdmin && onOpenVendorConsole && (
                  <button
                    type="button"
                    onClick={onOpenVendorConsole}
                    className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-xs font-semibold text-slate-800 cursor-pointer whitespace-nowrap"
                  >
                    <KeyRound className="w-3.5 h-3.5" />
                    <span>Vendor License Console</span>
                  </button>
                )}
              </div>

              {formError && (
                <div className="rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-xs text-red-800">
                  {formError}
                </div>
              )}

              {/* Pending Invitations */}
              {pendingInvites.length > 0 && (
                <div className="space-y-3">
                  <h3 className="text-sm font-semibold text-slate-900">
                    Pending Seat Invitations ({pendingInvites.length})
                  </h3>
                  <div className="divide-y divide-slate-200 border border-slate-200 rounded-lg">
                    {pendingInvites.map((inv) => (
                      <div
                        key={inv.membershipId}
                        className="p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                      >
                        <div>
                          <div className="text-sm font-semibold text-slate-900">
                            {inv.organizationName}
                          </div>
                          <div className="text-xs text-slate-500 mt-0.5 flex items-center gap-2">
                            <span>{ROLE_LABELS[inv.role]}</span>
                            <span aria-hidden="true">·</span>
                            <span className="font-mono">{inv.orgCode}</span>
                          </div>
                        </div>
                        <button
                          type="button"
                          disabled={acceptingId === inv.membershipId}
                          onClick={() => handleAcceptInvite(inv)}
                          className="inline-flex items-center justify-center gap-1.5 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white px-4 py-2 text-xs font-semibold cursor-pointer whitespace-nowrap disabled:opacity-60"
                        >
                          <CheckCircle2 className="w-3.5 h-3.5" />
                          <span>
                            {acceptingId === inv.membershipId ? 'Activating Seat...' : 'Accept Seat'}
                          </span>
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Create New Organization */}
              <form onSubmit={handleCreateOrg} className="space-y-4">
                <div>
                  <h3 className="text-sm font-semibold text-slate-900 flex items-center gap-2">
                    <Building2 className="w-4 h-4 text-slate-500" />
                    <span>Create New Organization Workspace</span>
                  </h3>
                  <p className="text-xs text-slate-500 mt-1">
                    Activate with a signed organization license key, or leave the key blank to start
                    with a 14-day, 5-seat trial.
                  </p>
                </div>

                <div>
                  <label
                    htmlFor="org-name-input"
                    className="block text-xs font-medium text-slate-700 mb-1"
                  >
                    Organization Name
                  </label>
                  <input
                    id="org-name-input"
                    type="text"
                    required
                    value={orgName}
                    onChange={(e) => setOrgName(e.target.value)}
                    placeholder="e.g. Grand Group FZC"
                    className="w-full min-h-[42px] rounded-lg border border-slate-300 bg-white px-3.5 py-2 text-sm text-slate-900 focus:border-[#1B365D] focus:outline-none"
                  />
                </div>

                <div>
                  <label
                    htmlFor="org-license-input"
                    className="block text-xs font-medium text-slate-700 mb-1"
                  >
                    Signed License Key <span className="text-slate-400 font-normal">(optional)</span>
                  </label>
                  <input
                    id="org-license-input"
                    type="text"
                    value={licenseKey}
                    onChange={(e) => setLicenseKey(e.target.value)}
                    placeholder="GGPRO-..."
                    className="w-full min-h-[42px] rounded-lg border border-slate-300 bg-white px-3.5 py-2 font-mono text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none"
                  />
                </div>

                <div className="pt-2 flex flex-wrap items-center justify-between gap-3">
                  <button
                    type="button"
                    onClick={onEnterLocalOfflineMode}
                    className="text-xs font-medium text-slate-500 hover:text-slate-800 cursor-pointer"
                  >
                    Or continue in standalone offline mode
                  </button>
                  <button
                    type="submit"
                    disabled={submitting}
                    className="inline-flex items-center gap-2 rounded-lg bg-[#1B365D] hover:bg-[#152a48] text-white px-5 py-2.5 text-xs font-semibold cursor-pointer whitespace-nowrap disabled:opacity-60"
                  >
                    <span>
                      {submitting
                        ? 'Creating Workspace...'
                        : licenseKey.trim()
                        ? 'Activate License & Create Workspace'
                        : 'Start 14-Day Trial Workspace'}
                    </span>
                    <ArrowRight className="w-3.5 h-3.5" />
                  </button>
                </div>
              </form>
            </div>
          )}
        </section>
      </main>
    </div>
  );
};
