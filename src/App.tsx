import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  createUserWithEmailAndPassword,
  onAuthStateChanged,
  signInWithEmailAndPassword,
  signInWithPopup,
  signOut,
  updateProfile,
  type User,
} from 'firebase/auth';
import {
  ArrowLeftRight,
  Building2,
  Calculator,
  CircleHelp,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Hash,
  KeyRound,
  LayoutDashboard,
  Lock,
  LogIn,
  LogOut,
  Menu,
  Monitor,
  Package,
  Plus,
  Radio,
  Receipt,
  Redo2,
  Send,
  Settings,
  ShieldAlert,
  ShoppingCart,
  Tablet,
  Trash2,
  Truck,
  Undo2,
  Unlock,
  Users,
  X,
} from 'lucide-react';
import { auth, googleAuthProvider } from './lib/firebase';
import { ensureUserProfile } from './lib/userProfile';

function friendlyAuthError(err: unknown): string {
  const code = typeof err === 'object' && err !== null && 'code' in err ? String((err as { code: unknown }).code) : '';
  switch (code) {
    case 'auth/invalid-email':
      return 'Enter a valid email address.';
    case 'auth/missing-password':
      return 'Enter your password.';
    case 'auth/weak-password':
      return 'Password must be at least 6 characters.';
    case 'auth/email-already-in-use':
      return 'An account with this email already exists. Sign in instead.';
    case 'auth/invalid-credential':
    case 'auth/wrong-password':
    case 'auth/user-not-found':
      return 'Email or password is incorrect.';
    case 'auth/too-many-requests':
      return 'Too many attempts. Please wait a few minutes and try again.';
    case 'auth/network-request-failed':
      return 'Network error. Check your connection and try again.';
    default:
      return err instanceof Error ? err.message : 'Sign-in failed. Please try again.';
  }
}
import {
  amendOrgDocument,
  ApiError,
  deleteOrgJob,
  fetchOrgAuditLog,
  fetchOrgWorkspace,
  fetchSessionBootstrap,
  issueOrgDocument,
  saveOrgJob,
  updateOrgSettings,
  voidOrgRegistryEntry,
} from './lib/apiClient';
import { AuthGateView } from './components/AuthGateView';
import { ExecutiveLiveView } from './components/ExecutiveLiveView';
import { OrgAdminView } from './components/OrgAdminView';
import { VendorLicenseConsoleView } from './components/VendorLicenseConsoleView';
import { useExecutiveStream } from './hooks/useExecutiveStream';
import {
  enqueueOfflineMutation,
  flushOfflineOutbox,
  loadOfflineOutbox,
  loadOrgCacheSnapshot,
  saveOrgCacheSnapshot,
} from './utils/offlineSync';
import {
  ROLE_LABELS,
  type AuthUserProfile,
  type OfflineLeaseToken,
  type OrganizationSummary,
  type PendingInvite,
} from './types/enterprise';
import { SupplierRfqView } from './components/SupplierRfqView';
import { applyPageSnapshot, undoPageKey } from './utils/pageUndo';
import { HelpView } from './components/HelpView';
import { helpTopicForTab } from './constants/helpTopics';
import { DOC_META, pageBarColor, resolveThemePalette } from './constants/themes';
import { ConfirmDialog } from './components/ConfirmDialog';
import { DocumentWorkspaceView } from './components/DocumentWorkspaceView';
import { ExcelStudioView } from './components/ExcelStudioView';
import { FilesHubView } from './components/FilesHubView';
import { InternalQuoteView } from './components/InternalQuoteView';
import { JobTrackerDashboard } from './components/JobTrackerDashboard';
import { OfflineIndicator } from './components/PWAInstallButton';
import { PriceCalcView } from './components/PriceCalcView';
import { RegistryView } from './components/RegistryView';
import { QuoteApplyOptions, SupplierQuoteImportModal } from './components/SupplierQuoteImportModal';
import {
  createBlankJob,
  createDefaultCompanyProfile,
  createDefaultWorkflow,
  formatDocNumber,
  getNextSequenceNumber,
  loadActiveJobId,
  loadCompanyProfile,
  loadDashboardPrefs,
  loadJobsWithHealth,
  loadRegistryWithHealth,
  loadThemeCustomization,
  loadThemeVariant,
  normalizeCompanyProfile,
  normalizeJob,
  mergeRegistryEntries,
  normalizeRegistryEntries,
  normalizeThemeCustomization,
  safeDeepClone,
  saveActiveJobId,
  saveCompanyProfile,
  saveDashboardPrefs,
  saveJobs,
  saveRegistry,
  saveThemeCustomization,
  saveThemeVariant,
  appendAuditLog,
  loadAuditLog,
  verifyAdminPassword,
} from './store/storage';
import { AmendmentDialog } from './components/AmendmentDialog';
import { auditLogToCsv, diffJobs, isJobLocked, purgeRevisionHistory } from './utils/amendments';
import { exportAuditLogXlsx } from './utils/excelEngine';
import { shareOrDownloadBlob } from './utils/platform';
import {
  CompanyProfile,
  ConfirmDialogRequest,
  CustomColumnDef,
  DashboardPreferences,
  DocPrefix,
  FreightConfig,
  InquiryType,
  JobControls,
  JobItemRow,
  JobWorkflowState,
  PlatformMode,
  ProcurementJob,
  RegistryEntry,
  SupplierQuoteData,
  SupplierRfqData,
  ThemeCustomization,
  ThemeVariantId,
  WorkspaceTab,
  AuditEntry,
  IssuedDocRecord,
  IssuedRevision,
} from './types/procurement';
import { calculateJob, formatMoney, formatPct, getResolvedEndUser } from './utils/calculations';
import { importWorkbookFromXlsx } from './utils/excelEngine';
import { detectDefaultPlatform, savePlatformMode } from './utils/platform';
import { ParsedSupplierQuote, parseSupplierPdfFile } from './utils/supplierQuoteParser';

const TAB_LABELS: Record<WorkspaceTab, string> = {
  EXECUTIVE: 'Executive Live Dashboard',
  TRACKER: 'Dashboard',
  SUPPLIER_RFQ: 'Supplier RFQ',
  INTERNAL_QUOTE: 'Internal Supplier Quote',
  PRICE_CALC: 'Price Calc (Internal)',
  QT: 'Quotation (Master Controls)',
  PI: 'Proforma Invoice (PI)',
  PO: 'Supplier PO (PO)',
  CI: 'Commercial Invoice (CI)',
  PL: 'Packing List (PL)',
  DN: 'Delivery Note (DN)',
  REGISTRY: 'Master Registry',
  FILES: 'Import & Export',
  STUDIO: 'Clients & Settings',
  ORG_ADMIN: 'Organization & Seats',
  VENDOR_LICENSES: 'Vendor License Console',
  HELP: 'Help & Troubleshooting',
};

export default function App() {
  // Single-pass startup load so jobs, activeJobId, and health recovery never diverge
  const [startup] = useState(() => {
    const jobsHealth = loadJobsWithHealth();
    const regHealth = loadRegistryWithHealth();
    const activeId = loadActiveJobId(jobsHealth.jobs);
    return {
      jobs: jobsHealth.jobs,
      registry: regHealth.entries,
      activeId,
      recovered: jobsHealth.recovered || regHealth.recovered,
    };
  });

  const [jobs, setJobs] = useState<ProcurementJob[]>(startup.jobs);
  const [activeJobId, setActiveJobId] = useState<string>(startup.activeId);
  const [registry, setRegistry] = useState<RegistryEntry[]>(startup.registry);
  const [themeId, setThemeId] = useState<ThemeVariantId>(() => loadThemeVariant());
  const [themeCustomization, setThemeCustomization] = useState<ThemeCustomization>(() =>
    loadThemeCustomization()
  );
  const [company, setCompany] = useState<CompanyProfile>(() => loadCompanyProfile());
  const [dashboardPrefs, setDashboardPrefs] = useState<DashboardPreferences>(() =>
    loadDashboardPrefs()
  );
  const [platformMode, setPlatformMode] = useState<PlatformMode>(() => detectDefaultPlatform());
  const [activeTab, setActiveTab] = useState<WorkspaceTab>('TRACKER');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [toast, setToast] = useState<string | null>(
    startup.recovered
      ? 'Recovered damaged local storage automatically. Your workspace is safe to use.'
      : null
  );
  const [confirmDialog, setConfirmDialog] = useState<ConfirmDialogRequest | null>(null);
  const [pendingSupplierQuote, setPendingSupplierQuote] = useState<ParsedSupplierQuote | null>(null);
  // Issued-document amendment: one document is unlocked at a time; the snapshot allows discard.
  const [amendment, setAmendment] = useState<{ jobId: string; prefix: DocPrefix; snapshot: ProcurementJob } | null>(null);
  // Help: remember the page the user came from, so Back returns to it
  const [helpReturnTab, setHelpReturnTab] = useState<WorkspaceTab>('TRACKER');
  const [helpTopicId, setHelpTopicId] = useState('overview');
  const openHelp = () => {
    if (activeTab !== 'HELP') setHelpReturnTab(activeTab);
    setHelpTopicId(helpTopicForTab(activeTab === 'HELP' ? helpReturnTab : activeTab));
    setActiveTab('HELP');
    setSidebarOpen(false);
  };
  const [amendDialog, setAmendDialog] = useState<{ mode: "unlock" | "save"; prefix: DocPrefix } | null>(null);
  const quickImportRef = useRef<HTMLInputElement | null>(null);

  const activeJob = useMemo(
    () => jobs.find((j) => j.id === activeJobId) || jobs[0],
    [jobs, activeJobId]
  );

  const calc = useMemo(() => calculateJob(activeJob), [activeJob]);
  const nextSeq = useMemo(() => getNextSequenceNumber(registry, jobs), [registry, jobs]);
  const theme = useMemo(
    () => resolveThemePalette(themeId, themeCustomization),
    [themeId, themeCustomization]
  );

  // ==================== MULTI-USER AUTH & ORGANIZATION STATE ====================
  const [authReady, setAuthReady] = useState(false);
  const [authLoading, setAuthLoading] = useState(false);
  const [authError, setAuthError] = useState<string | null>(null);
  const [firebaseUser, setFirebaseUser] = useState<User | null>(null);
  const [authUser, setAuthUser] = useState<AuthUserProfile | null>(null);
  const bootstrapInFlightRef = useRef<Promise<void> | null>(null);
  const [organizations, setOrganizations] = useState<OrganizationSummary[]>([]);
  const [pendingInvites, setPendingInvites] = useState<PendingInvite[]>([]);
  const [activeOrgId, setActiveOrgId] = useState<string | null>(null);
  const [localStandaloneMode, setLocalStandaloneMode] = useState(false);
  const [onboardingVendorOpen, setOnboardingVendorOpen] = useState(false);
  const [offlineLease, setOfflineLease] = useState<OfflineLeaseToken | null>(null);
  const [pendingOutboxCount, setPendingOutboxCount] = useState(0);
  const [syncConflict, setSyncConflict] = useState<{
    jobId: string;
    localJob: ProcurementJob;
    serverJob: ProcurementJob;
    message: string;
  } | null>(null);
  const [executiveSafeMode, setExecutiveSafeMode] = useState(false);
  const [workspaceLoading, setWorkspaceLoading] = useState(false);

  const activeOrg = useMemo(() => {
    const list = Array.isArray(organizations) ? organizations : [];
    return list.find((o) => o && o.id === activeOrgId) || null;
  }, [organizations, activeOrgId]);
  const isCloudMode = Boolean(!localStandaloneMode && authUser && activeOrg);
  const isReadOnlyRole = Boolean(
    isCloudMode &&
      activeOrg &&
      (activeOrg.role === 'VIEWER' ||
        activeOrg.licenseExpired ||
        ((activeOrg.role === 'CEO_EXECUTIVE' || activeOrg.role === 'ORG_ADMIN') &&
          executiveSafeMode))
  );

  // Refs to prevent remote stream updates from re-triggering saves
  const suppressCloudSaveRef = useRef(false);
  const jobSaveTimersRef = useRef<Map<string, ReturnType<typeof setTimeout>>>(new Map());

  const showToast = useCallback((msg: string) => {
    setToast(msg);
  }, []);

  const loadOrganizationWorkspace = useCallback(
    async (org: OrganizationSummary) => {
      setWorkspaceLoading(true);
      try {
        const snap = await fetchOrgWorkspace(org.id);
        suppressCloudSaveRef.current = true;
        saveOrgCacheSnapshot(org.id, snap);
        setOfflineLease(snap.offlineLease);
        setOrganizations((prev) =>
          (Array.isArray(prev) ? prev : []).map((o) =>
            o && o.id === snap.organization.id ? snap.organization : o
          )
        );
        const hydratedJobs =
          snap.jobs.length > 0
            ? snap.jobs.map((j) => normalizeJob(j))
            : [createBlankJob('Procurement Job #1', snap.companyProfile)];
        setJobs(hydratedJobs);
        setActiveJobId((prevId) =>
          hydratedJobs.some((j) => j.id === prevId) ? prevId : hydratedJobs[0].id
        );
        setRegistry(normalizeRegistryEntries(snap.registry));
        setCompany(normalizeCompanyProfile(snap.companyProfile));
        setThemeId(snap.themeVariant);
        setThemeCustomization(normalizeThemeCustomization(snap.themeCustomization));
        setDashboardPrefs(snap.dashboardPrefs);
        setPendingOutboxCount(loadOfflineOutbox(org.id).length);

        if (snap.organization.role === 'CEO_EXECUTIVE') {
          setExecutiveSafeMode(true);
          setActiveTab('EXECUTIVE');
        } else if (snap.organization.licenseExpired && snap.organization.role === 'ORG_ADMIN') {
          setActiveTab('ORG_ADMIN');
        }
        setTimeout(() => {
          suppressCloudSaveRef.current = false;
        }, 120);
      } catch (err) {
        // Offline resilience: fall back to validated 7-day cached snapshot if available
        const cached = loadOrgCacheSnapshot(org.id);
        if (cached) {
          suppressCloudSaveRef.current = true;
          setOfflineLease(cached.offlineLease);
          const hydratedJobs =
            cached.jobs.length > 0
              ? cached.jobs.map((j) => normalizeJob(j))
              : [createBlankJob('Procurement Job #1', cached.companyProfile)];
          setJobs(hydratedJobs);
          setActiveJobId(hydratedJobs[0].id);
          setRegistry(normalizeRegistryEntries(cached.registry));
          setCompany(normalizeCompanyProfile(cached.companyProfile));
          setThemeId(cached.themeVariant);
          setThemeCustomization(normalizeThemeCustomization(cached.themeCustomization));
          setDashboardPrefs(cached.dashboardPrefs);
          setPendingOutboxCount(loadOfflineOutbox(org.id).length);
          setToast(
            `Offline cache loaded (signed lease valid until ${cached.offlineLease.expiresAt.slice(0, 10)}). Draft edits will queue and sync on reconnect.`
          );
          setTimeout(() => {
            suppressCloudSaveRef.current = false;
          }, 120);
        } else {
          setToast(
            err instanceof Error
              ? err.message
              : 'Could not load organization workspace from server.'
          );
        }
      } finally {
        setWorkspaceLoading(false);
      }
    },
    []
  );

  const formatAuthError = (err: unknown): string => {
    const code =
      typeof err === 'object' && err !== null && 'code' in err
        ? String((err as { code?: unknown }).code || '')
        : '';
    switch (code) {
      case 'auth/popup-blocked':
        return 'Your browser blocked the Google Sign-In popup. Please allow popups for this site and tap Continue with Google again.';
      case 'auth/popup-closed-by-user':
      case 'auth/cancelled-popup-request':
        return 'Google Sign-In was closed before completing. Tap Continue with Google to try again.';
      case 'auth/unauthorized-domain':
        return `This domain (${window.location.hostname}) is not authorized in Firebase Authentication yet. Add it under Firebase Console → Authentication → Settings → Authorized domains.`;
      case 'auth/operation-not-allowed':
        return 'Google Sign-In is not enabled in Firebase Authentication for this project.';
      case 'auth/network-request-failed':
        return 'Network error contacting Google Sign-In. Check your internet connection and try again.';
      default:
        return err instanceof Error ? err.message : 'Could not verify your workspace session.';
    }
  };

  const bootstrapSession = useCallback(
    async (explicitUser?: User | null) => {
      if (bootstrapInFlightRef.current) {
        return bootstrapInFlightRef.current;
      }

      const task = (async () => {
        const fbUser = explicitUser ?? auth.currentUser;
        if (!fbUser) {
          setAuthUser(null);
          setOrganizations([]);
          setPendingInvites([]);
          setActiveOrgId(null);
          setAuthLoading(false);
          setAuthReady(true);
          return;
        }

        setAuthLoading(true);
        setAuthError(null);
        try {
          const freshIdToken = await fbUser.getIdToken(true);
          const session = await fetchSessionBootstrap(freshIdToken);
          if (!session || !session.user || !Array.isArray(session.organizations)) {
            throw new Error('Workspace session response was incomplete. Please sign in again.');
          }
          const nextOrgs = session.organizations.filter(
            (o): o is OrganizationSummary => Boolean(o && typeof o.id === 'string')
          );
          const nextInvites = Array.isArray(session.pendingInvites) ? session.pendingInvites : [];
          setAuthUser(session.user);
          setOrganizations(nextOrgs);
          setPendingInvites(nextInvites);
          if (nextOrgs.length > 0) {
            const chosen = nextOrgs[0];
            setActiveOrgId(chosen.id);
            setLocalStandaloneMode(false);
            await loadOrganizationWorkspace(chosen);
          } else {
            setActiveOrgId(null);
          }
        } catch (err) {
          setAuthError(formatAuthError(err));
        } finally {
          setAuthLoading(false);
          setAuthReady(true);
        }
      })();

      bootstrapInFlightRef.current = task;
      try {
        await task;
      } finally {
        if (bootstrapInFlightRef.current === task) {
          bootstrapInFlightRef.current = null;
        }
      }
    },
    [loadOrganizationWorkspace]
  );

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (fbUser) => {
      setFirebaseUser(fbUser);
      if (fbUser) {
        await bootstrapSession(fbUser);
      } else {
        setAuthUser(null);
        setOrganizations([]);
        setPendingInvites([]);
        setActiveOrgId(null);
        setAuthLoading(false);
        setAuthReady(true);
      }
    });
    return () => unsub();
  }, [bootstrapSession]);

  const handleGoogleSignIn = async () => {
    setAuthLoading(true);
    setAuthError(null);
    try {
      const cred = await signInWithPopup(auth, googleAuthProvider);
      if (cred.user) {
        setFirebaseUser(cred.user);
        await ensureUserProfile(cred.user).catch((err) =>
          console.warn('Could not write Firestore user profile:', err)
        );
        await bootstrapSession(cred.user);
      }
    } catch (err) {
      setAuthError(formatAuthError(err));
    } finally {
      setAuthLoading(false);
    }
  };

  const handleEmailSignIn = async (email: string, password: string) => {
    setAuthLoading(true);
    setAuthError(null);
    try {
      const cred = await signInWithEmailAndPassword(auth, email.trim(), password);
      await ensureUserProfile(cred.user).catch((err) =>
        console.warn('Could not write Firestore user profile:', err)
      );
    } catch (err) {
      setAuthError(friendlyAuthError(err));
      setAuthLoading(false);
    }
  };

  const handleEmailSignUp = async (name: string, email: string, password: string) => {
    setAuthLoading(true);
    setAuthError(null);
    try {
      const cred = await createUserWithEmailAndPassword(auth, email.trim(), password);
      const displayName = name.trim();
      if (displayName) {
        await updateProfile(cred.user, { displayName });
      }
      await ensureUserProfile(cred.user, displayName || undefined).catch((err) =>
        console.warn('Could not write Firestore user profile:', err)
      );
    } catch (err) {
      setAuthError(friendlyAuthError(err));
      setAuthLoading(false);
    }
  };

  const handleRetrySession = async () => {
    const current = auth.currentUser ?? firebaseUser;
    if (current) {
      await bootstrapSession(current);
    } else {
      await handleGoogleSignIn();
    }
  };

  const handleSignOut = async () => {
    try {
      await signOut(auth);
      setFirebaseUser(null);
      setActiveOrgId(null);
      setAuthUser(null);
      setAuthError(null);
      setOrganizations([]);
      setPendingInvites([]);
      setLocalStandaloneMode(false);
      setOnboardingVendorOpen(false);
    } catch {
      showToast('Could not sign out cleanly.');
    }
  };

  const handleFlushOfflineOutbox = useCallback(async () => {
    if (!activeOrgId || !isCloudMode) return;
    const result = await flushOfflineOutbox(activeOrgId);
    setPendingOutboxCount(result.remaining.length);
    if (result.conflicts.length > 0) {
      setSyncConflict(result.conflicts[0]);
      showToast('A sync conflict was detected. Review the banner to choose which version to keep.');
    } else if (result.syncedCount > 0) {
      showToast(`Synced ${result.syncedCount} queued offline change(s) to the cloud workspace.`);
    }
  }, [activeOrgId, isCloudMode, showToast]);

  useEffect(() => {
    if (!isCloudMode || !activeOrgId) return;
    const onOnline = () => {
      handleFlushOfflineOutbox();
    };
    window.addEventListener('online', onOnline);
    return () => window.removeEventListener('online', onOnline);
  }, [isCloudMode, activeOrgId, handleFlushOfflineOutbox]);

  const queueCloudJobSync = useCallback(
    (jobToSync: ProcurementJob, options?: { immediate?: boolean; auditReason?: string; forceOverwrite?: boolean }) => {
      if (!isCloudMode || !activeOrgId || suppressCloudSaveRef.current) return;

      const execute = async () => {
        try {
          const res = await saveOrgJob(activeOrgId, jobToSync, {
            baseVersion: jobToSync.version,
            forceOverwrite: options?.forceOverwrite,
            auditReason: options?.auditReason,
          });
          setJobs((prev) =>
            prev.map((j) => (j.id === res.job.id ? { ...j, version: res.version } : j))
          );
        } catch (err) {
          if (err instanceof ApiError && err.status === 409) {
            const data = err.data as { serverJob?: ProcurementJob } | undefined;
            if (data?.serverJob) {
              setSyncConflict({
                jobId: jobToSync.id,
                localJob: jobToSync,
                serverJob: normalizeJob(data.serverJob),
                message: err.message,
              });
              return;
            }
          }
          const updatedQueue = enqueueOfflineMutation(activeOrgId, {
            kind: 'upsert_job',
            jobId: jobToSync.id,
            payload: jobToSync,
          });
          setPendingOutboxCount(updatedQueue.length);
        }
      };

      const existingTimer = jobSaveTimersRef.current.get(jobToSync.id);
      if (existingTimer) clearTimeout(existingTimer);

      if (options?.immediate) {
        jobSaveTimersRef.current.delete(jobToSync.id);
        void execute();
      } else {
        const t = setTimeout(() => {
          jobSaveTimersRef.current.delete(jobToSync.id);
          void execute();
        }, 650);
        jobSaveTimersRef.current.set(jobToSync.id, t);
      }
    },
    [isCloudMode, activeOrgId]
  );

  // Real-time SSE Live Stream for CEO Executive Dashboard & Multi-User Workspace Sync
  const executiveStream = useExecutiveStream({
    orgId: activeOrgId,
    enabled: Boolean(isCloudMode && activeOrg && !activeOrg.licenseExpired),
    onRemoteJobUpdated: (remoteJob, updatedByUid) => {
      if (updatedByUid && authUser && updatedByUid === authUser.uid) return;
      const norm = normalizeJob(remoteJob);
      setJobs((prev) => {
        const exists = prev.some((j) => j.id === norm.id);
        return exists ? prev.map((j) => (j.id === norm.id ? norm : j)) : [norm, ...prev];
      });
    },
    onRemoteJobDeleted: (deletedJobId, deletedByUid) => {
      if (deletedByUid && authUser && deletedByUid === authUser.uid) return;
      setJobs((prev) => {
        const filtered = prev.filter((j) => j.id !== deletedJobId);
        return filtered.length > 0 ? filtered : prev;
      });
    },
    onRemoteDocIssued: (remoteJob, regEntry) => {
      const norm = normalizeJob(remoteJob);
      setJobs((prev) => prev.map((j) => (j.id === norm.id ? norm : j)));
      setRegistry((prev) => mergeRegistryEntries(prev, [regEntry]));
    },
    onRemoteRegistryVoided: (docNumber, voidReason) => {
      setRegistry((prev) =>
        prev.map((e) => (e.docNumber === docNumber ? { ...e, status: 'VOID', voidReason } : e))
      );
    },
    onSeatRevoked: (reason) => {
      showToast(reason);
      setActiveOrgId(null);
      void bootstrapSession();
    },
  });

  useEffect(() => {
    if (localStandaloneMode) {
      if (!saveJobs(jobs)) {
        setToast('Local storage is nearly full. Export a JSON Backup in Import & Export to safeguard your jobs.');
      }
    }
  }, [jobs, localStandaloneMode]);

  useEffect(() => {
    if (localStandaloneMode) {
      if (!saveRegistry(registry)) {
        setToast('Local storage is nearly full. Export Master_Registry.xlsx or JSON Backup to protect records.');
      }
    }
  }, [registry, localStandaloneMode]);

  useEffect(() => {
    const purged = purgeRevisionHistory(jobs);
    if (purged !== jobs) setJobs(purged);
  }, [jobs]);

  useEffect(() => {
    if (localStandaloneMode) saveThemeVariant(themeId);
    else if (isCloudMode && activeOrgId && !suppressCloudSaveRef.current) {
      updateOrgSettings(activeOrgId, { themeVariant: themeId }).catch(() => {});
    }
  }, [themeId, localStandaloneMode, isCloudMode, activeOrgId]);

  useEffect(() => {
    if (localStandaloneMode) saveThemeCustomization(themeCustomization);
    else if (isCloudMode && activeOrgId && !suppressCloudSaveRef.current) {
      updateOrgSettings(activeOrgId, { themeCustomization }).catch(() => {});
    }
  }, [themeCustomization, localStandaloneMode, isCloudMode, activeOrgId]);

  useEffect(() => {
    if (localStandaloneMode) {
      if (!saveCompanyProfile(company)) {
        setToast('Could not save company settings (logo image may be too large for browser storage).');
      }
    }
  }, [company, localStandaloneMode]);

  useEffect(() => savePlatformMode(platformMode), [platformMode]);
  useEffect(() => {
    if (localStandaloneMode) saveDashboardPrefs(dashboardPrefs);
    else if (isCloudMode && activeOrgId && !suppressCloudSaveRef.current) {
      updateOrgSettings(activeOrgId, { dashboardPrefs }).catch(() => {});
    }
  }, [dashboardPrefs, localStandaloneMode, isCloudMode, activeOrgId]);

  useEffect(() => {
    if (activeJob?.id) saveActiveJobId(activeJob.id);
  }, [activeJob?.id]);

  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => setToast(null), 4200);
    return () => clearTimeout(timer);
  }, [toast]);

  const UNDO_LIMIT = 15;
  const UNDO_COALESCE_MS = 1500;
  const undoHistoryRef = useRef<Record<string, { past: ProcurementJob[]; future: ProcurementJob[]; lastEditAt: number }>>({});
  const [, setHistoryVersion] = useState(0);
  const historyFor = (key: string) => {
    if (!undoHistoryRef.current[key]) undoHistoryRef.current[key] = { past: [], future: [], lastEditAt: 0 };
    return undoHistoryRef.current[key];
  };

  const currentUndoKey = () => `${activeJob.id}|${undoPageKey(activeTab)}`;

  const mutateSystemJob = (updater: (job: ProcurementJob) => ProcurementJob, options?: { skipCloudSync?: boolean; auditReason?: string }) => {
    if (isReadOnlyRole) {
      showToast(
        executiveSafeMode
          ? 'Safe Read-Only Mode is ON. Switch to Full Edit Access in the toolbar to modify jobs.'
          : 'Your current seat or license state is read-only.'
      );
      return;
    }
    const h = historyFor(currentUndoKey());
    const now = Date.now();
    if (now - h.lastEditAt > UNDO_COALESCE_MS) {
      h.past.push(activeJob);
      if (h.past.length > UNDO_LIMIT) h.past.shift();
      h.future = [];
      setHistoryVersion((v) => v + 1);
    }
    h.lastEditAt = now;
    let updatedRef: ProcurementJob | null = null;
    setJobs((prev) =>
      prev.map((j) => {
        if (j.id !== activeJob.id) return j;
        updatedRef = { ...updater(j), updatedAt: new Date().toISOString() };
        return updatedRef;
      })
    );
    if (updatedRef && !options?.skipCloudSync) {
      queueCloudJobSync(updatedRef, { auditReason: options?.auditReason });
    }
  };

  const isEditingLocked = () => isJobLocked(activeJob) && amendment?.jobId !== activeJob.id;

  // Undo / redo restore the job's editable content. Issued-document records stay as they are.
  const restoreJobSnapshot = (snapshot: ProcurementJob) => {
    const tab = activeTab;
    setJobs((prev) =>
      prev.map((j) =>
        j.id === activeJob.id
          ? { ...applyPageSnapshot(j, snapshot, tab), updatedAt: new Date().toISOString() }
          : j
      )
    );
  };

  const handleUndo = () => {
    if (isEditingLocked()) return showToast('Issued document is locked. Use Unlock on the document to make changes.');
    const h = historyFor(currentUndoKey());
    const previous = h.past.pop();
    if (!previous) return showToast('Nothing to undo for this job.');
    h.future.push(activeJob);
    h.lastEditAt = 0;
    restoreJobSnapshot(previous);
    setHistoryVersion((v) => v + 1);
  };

  const handleRedo = () => {
    if (isEditingLocked()) return showToast('Issued document is locked. Use Unlock on the document to make changes.');
    const h = historyFor(currentUndoKey());
    const next = h.future.pop();
    if (!next) return showToast('Nothing to redo for this job.');
    h.past.push(activeJob);
    h.lastEditAt = 0;
    restoreJobSnapshot(next);
    setHistoryVersion((v) => v + 1);
  };

  // Ctrl+Z / Ctrl+Y (Cmd on Mac). Inside text fields the browser keeps its own undo.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      const target = e.target as HTMLElement | null;
      if (target && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName))) return;
      const key = e.key.toLowerCase();
      if (key === 'z' && !e.shiftKey) {
        e.preventDefault();
        handleUndo();
      } else if (key === 'y' || (key === 'z' && e.shiftKey)) {
        e.preventDefault();
        handleRedo();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  // User edits: blocked on a locked job unless this job is open for amendment.
  const mutateActiveJob = (updater: (job: ProcurementJob) => ProcurementJob) => {
    const amending = amendment?.jobId === activeJob.id;
    if (isJobLocked(activeJob) && !amending) {
      showToast("Issued document is locked. Use Unlock on the document to make changes.");
      return;
    }
    mutateSystemJob(updater);
  };

  const handleUpdateJobName = (name: string) => {
    mutateActiveJob((j) => ({ ...j, jobName: name }));
  };

  const handleUpdateInquiryType = (inquiryType: InquiryType) => {
    mutateActiveJob((j) => ({ ...j, inquiryType }));
  };

  const handleUpdateControls = (patch: Partial<JobControls>) => {
    mutateActiveJob((j) => ({
      ...j,
      controls: { ...j.controls, ...patch },
    }));
  };

  const handleUpdateSupplierQuote = (patch: Partial<SupplierQuoteData>) => {
    mutateActiveJob((j) => {
      const nextSq = { ...j.supplierQuote, ...patch };
      const nextControls = { ...j.controls };
      if (patch.supplierName !== undefined) nextControls.supplierName = patch.supplierName;
      if (patch.supplierAddress !== undefined) nextControls.supplierAddress = patch.supplierAddress;
      if (patch.supplierQuoteRef !== undefined) nextControls.supplierRef = patch.supplierQuoteRef;
      return {
        ...j,
        supplierQuote: nextSq,
        controls: nextControls,
      };
    });
  };

  const handleUpdateWorkflow = (patch: Partial<JobWorkflowState>) => {
    mutateSystemJob((j) => ({
      ...j,
      workflow: { ...j.workflow, ...patch },
    }));
  };

  const handleUpdateDashboardPrefs = (patch: Partial<DashboardPreferences>) => {
    setDashboardPrefs((prev) => ({ ...prev, ...patch }));
  };

  const handleSelectJobAndTab = (jobId: string, tab: WorkspaceTab) => {
    setActiveJobId(jobId);
    setActiveTab(tab);
    setSidebarOpen(false);
  };

  const handleUpdateFreight = (patch: Partial<FreightConfig>) => {
    mutateActiveJob((j) => ({
      ...j,
      freight: { ...j.freight, ...patch },
    }));
  };

  const handleUpdateCustomColumns = (
    cols: CustomColumnDef[],
    nextColumnOrder?: string[]
  ) => {
    mutateActiveJob((j) => ({
      ...j,
      customColumns: cols,
      columnOrder: nextColumnOrder !== undefined ? nextColumnOrder : j.columnOrder,
    }));
  };

  const handleUpdateColumnOrder = (order: string[] | undefined) => {
    mutateActiveJob((j) => ({
      ...j,
      columnOrder: order,
    }));
  };

  const handleUpdateRfq = (patch: Partial<SupplierRfqData>) => {
    mutateActiveJob((j) => ({
      ...j,
      supplierRfq: {
        rfqNumber: '',
        supplierName: '',
        supplierAddress: '',
        supplierContact: '',
        supplierEmail: '',
        responseDueDate: '',
        showEndUser: false,
        sourceName: '',
        currency: '',
        incoterm: '',
        destination: '',
        ...j.supplierRfq,
        ...patch,
      },
    }));
  };

  // Uploaded end-user RFQ lines go into empty rows only, in order. Filled rows are never overwritten.
  const handleFillEmptyRows = (items: { requestedDescription: string; uom: string; qty: number | null }[]) => {
    mutateActiveJob((j) => {
      const queue = [...items];
      return {
        ...j,
        items: j.items.map((it) => {
          if (queue.length === 0 || it.requestedDescription || it.offeredDescription) return it;
          const next = queue.shift()!;
          return { ...it, requestedDescription: next.requestedDescription, offeredDescription: next.requestedDescription, uom: next.uom, qty: next.qty };
        }),
      };
    });
  };

  const handleUpdateItem = (rowNumber: number, patch: Partial<JobItemRow>) => {
    mutateActiveJob((j) => ({
      ...j,
      items: j.items.map((item) =>
        item.rowNumber === rowNumber ? { ...item, ...patch } : item
      ),
    }));
  };

  const handleCreateNewJob = () => {
    if (isReadOnlyRole) {
      showToast('Switch off Safe Read-Only Mode to create a new job.');
      return;
    }
    const num = jobs.length + 1;
    const fresh = createBlankJob(`Procurement Job #${num}`, company);
    if (authUser) {
      fresh.preparedBy = authUser.displayName || authUser.email;
    }
    setJobs((prev) => [fresh, ...prev]);
    setActiveJobId(fresh.id);
    setActiveTab('QT');
    setSidebarOpen(false);
    queueCloudJobSync(fresh, { immediate: true, auditReason: 'Created blank job' });
    showToast(`Created fresh job: ${fresh.jobName}`);
  };

  const handleDuplicateAsNewJob = () => {
    if (isReadOnlyRole) {
      showToast('Switch off Safe Read-Only Mode to duplicate a job.');
      return;
    }
    const copy: ProcurementJob = {
      ...safeDeepClone(activeJob),
      id: `job-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      jobName: `${activeJob.jobName} (Revision)`,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      preparedBy: authUser?.displayName || authUser?.email || activeJob.preparedBy,
      jobSequence: undefined,
      jobSequenceYear: undefined,
      version: 1,
      workflow: createDefaultWorkflow(),
      issuedDocs: {},
    };
    setJobs((prev) => [copy, ...prev]);
    setActiveJobId(copy.id);
    setActiveTab('QT');
    queueCloudJobSync(copy, { immediate: true, auditReason: 'Duplicated as revision job' });
    showToast(`Created unlocked revision job: ${copy.jobName} (will receive a new sequence when issued)`);
  };

  const handleDeleteCurrentJob = () => {
    if (isReadOnlyRole) {
      showToast('Switch off Safe Read-Only Mode to delete a job.');
      return;
    }
    if (jobs.length <= 1) return;
    const target = activeJob;
    setConfirmDialog({
      title: `Delete "${target.jobName}"?`,
      message: isCloudMode
        ? 'This permanently removes the job from your organization workspace. Claimed document numbers remain recorded in the Master Registry.'
        : 'This permanently removes the job from your local workspace. Claimed document numbers remain recorded in the Master Registry.',
      confirmLabel: 'Delete Job',
      destructive: true,
      onConfirm: () => {
        setJobs((prev) => {
          const remaining = prev.filter((j) => j.id !== target.id);
          if (remaining.length === 0) {
            const fresh = createBlankJob('Procurement Job #1', company);
            setActiveJobId(fresh.id);
            return [fresh];
          }
          setActiveJobId(remaining[0].id);
          return remaining;
        });
        if (isCloudMode && activeOrgId) {
          deleteOrgJob(activeOrgId, target.id).catch(() => {
            const q = enqueueOfflineMutation(activeOrgId, {
              kind: 'delete_job',
              jobId: target.id,
              payload: null,
            });
            setPendingOutboxCount(q.length);
          });
        }
        showToast(`Removed job "${target.jobName}".`);
      },
    });
  };

  const handleIssueDocument = (prefix: DocPrefix) => {
    if (isReadOnlyRole) {
      showToast('Switch off Safe Read-Only Mode to issue trade documents.');
      return;
    }

    if (isCloudMode && activeOrgId) {
      if (!navigator.onLine) {
        showToast(
          'Connect to the server before issuing a document so a unique Master Registry sequence number can be locked safely.'
        );
        return;
      }
      issueOrgDocument(activeOrgId, activeJob, prefix)
        .then((res) => {
          const norm = normalizeJob(res.job);
          setJobs((prev) => prev.map((j) => (j.id === norm.id ? norm : j)));
          setRegistry((prev) => mergeRegistryEntries(prev, [res.registryEntry]));
          showToast(
            `Issued ${res.docNumber} (Atomic Org Seq #${String(res.sequence).padStart(3, '0')}) — streamed live to executive dashboard.`
          );
        })
        .catch((err) => {
          showToast(err instanceof Error ? err.message : 'Could not issue document on server.');
        });
      return;
    }

    // Standalone offline local mode fallback
    const hasAssignedSeq =
      typeof activeJob.jobSequence === 'number' && activeJob.jobSequence > 0;
    const seq = hasAssignedSeq
      ? activeJob.jobSequence!
      : getNextSequenceNumber(registry, jobs);
    const year =
      hasAssignedSeq && activeJob.jobSequenceYear
        ? activeJob.jobSequenceYear
        : new Date().getFullYear();
    const docNumber = formatDocNumber(prefix, seq, year);
    const now = new Date().toISOString();

    const counterpartyRole = prefix === 'PO' ? 'SUPPLIER' : 'CLIENT';
    const partyName =
      prefix === 'PO'
        ? activeJob.controls.supplierName || 'Supplier'
        : getResolvedEndUser(activeJob) || 'Client';
    const docAmount =
      prefix === 'PO'
        ? calc.totalSupplierCost
        : prefix === 'PL' || prefix === 'DN'
        ? 0
        : calc.grandTotal;

    const newRegEntry: RegistryEntry = {
      sequence: seq,
      docNumber,
      prefix,
      year,
      jobId: activeJob.id,
      jobName: activeJob.jobName,
      customer: partyName,
      counterpartyRole,
      rfqRef: activeJob.controls.rfqRef || '',
      currency: activeJob.controls.currency,
      amount: docAmount,
      issuedBy: activeJob.preparedBy || 'Procurement Desk',
      issuedAt: now,
      status: 'ISSUED',
    };

    setRegistry((prev) => [...prev, newRegEntry]);
    mutateSystemJob((j) => {
      let nextControls = j.controls;
      let nextItems = j.items;
      const nextWorkflow = j.workflow;

      if (prefix === 'PL') {
        nextControls = {
          ...j.controls,
          deliveryDate: j.controls.deliveryDate.trim() || j.controls.shipDate || now.slice(0, 10),
        };
        nextItems = j.items.map((item, idx) => {
          const comp = calc.rows[idx];
          if (!comp || comp.isEmpty) return item;
          return {
            ...item,
            grossWtKg: item.grossWtKg ?? comp.grossKgCalc,
          };
        });
      }

      return {
        ...j,
        jobSequence: seq,
        jobSequenceYear: year,
        controls: nextControls,
        items: nextItems,
        workflow: nextWorkflow,
        issuedDocs: {
          ...j.issuedDocs,
          [prefix]: {
            prefix,
            docNumber,
            sequence: seq,
            issuedAt: now,
            issuedBy: j.preparedBy || 'Procurement Desk',
            frozenSubtotal: calc.goodsSubtotal,
            frozenVat: calc.vatAmount,
            frozenFreight: calc.customerFreight,
            frozenGrandTotal: calc.grandTotal,
          },
        },
      };
    });

    if (prefix === 'PL') {
      showToast(
        `Issued ${docNumber} (Job Seq #${String(seq).padStart(3, '0')}). Delivery Note populated and ready for manual issue.`
      );
    } else if (!hasAssignedSeq) {
      showToast(
        `Issued ${docNumber} — Job Seq #${String(seq).padStart(3, '0')} locked for all documents in this job.`
      );
    } else {
      showToast(`Issued ${docNumber} sharing Job Seq #${String(seq).padStart(3, '0')}.`);
    }
  };

  const handleVoidRegistryEntry = (docNumber: string, reason: string) => {
    if (isCloudMode && activeOrgId) {
      voidOrgRegistryEntry(activeOrgId, docNumber, reason)
        .then((res) => {
          setRegistry((prev) =>
            prev.map((e) => (e.docNumber === docNumber ? res.entry : e))
          );
          setJobs((prev) =>
            prev.map((j) => {
              const nextIssued = { ...j.issuedDocs };
              let changed = false;
              (Object.keys(nextIssued) as DocPrefix[]).forEach((p) => {
                if (nextIssued[p]?.docNumber === docNumber) {
                  nextIssued[p] = { ...nextIssued[p]!, isVoided: true };
                  changed = true;
                }
              });
              return changed ? { ...j, issuedDocs: nextIssued } : j;
            })
          );
          showToast(`Marked ${docNumber} as VOID across organization Master Registry.`);
        })
        .catch((err) => {
          showToast(err instanceof Error ? err.message : 'Could not void registry entry.');
        });
      return;
    }

    setRegistry((prev) =>
      prev.map((e) =>
        e.docNumber === docNumber ? { ...e, status: 'VOID', voidReason: reason } : e
      )
    );
    setJobs((prev) =>
      prev.map((j) => {
        const nextIssued = { ...j.issuedDocs };
        let changed = false;
        (Object.keys(nextIssued) as DocPrefix[]).forEach((p) => {
          if (nextIssued[p]?.docNumber === docNumber) {
            nextIssued[p] = { ...nextIssued[p]!, isVoided: true };
            changed = true;
          }
        });
        return changed ? { ...j, issuedDocs: nextIssued } : j;
      })
    );
    showToast(`Marked ${docNumber} as VOID. Document number permanently retired.`);
  };

  const handleSyncRegistry = (incoming: RegistryEntry[]) => {
    setRegistry((prev) => mergeRegistryEntries(prev, incoming));
  };

  const handleImportJob = (imported: ProcurementJob) => {
    const normalized = normalizeJob(imported);
    setJobs((prev) => [normalized, ...prev]);
    setActiveJobId(normalized.id);
    setActiveTab('QT');
    queueCloudJobSync(normalized, { immediate: true, auditReason: 'Imported workbook job' });
  };

  const handleApplySupplierQuote = (
    quote: ParsedSupplierQuote,
    options: QuoteApplyOptions
  ) => {
    const selectedItems = quote.items.filter((it) => it.selected).slice(0, 20);
    if (selectedItems.length === 0) return;

    const buildUpdatedJob = (base: ProcurementJob): ProcurementJob => {
      const existingCols = base.customColumns || [];
      let nextCols = [...existingCols];
      let codeColId: string | null = null;

      if (options.addCodeColumn) {
        const matchCol = existingCols.find(
          (c) => c.label.toLowerCase() === options.codeColumnLabel.toLowerCase()
        );
        if (matchCol) {
          codeColId = matchCol.id;
        } else if (nextCols.length < 4) {
          codeColId = `col_${options.codeColumnLabel.toLowerCase().replace(/[^a-z0-9]+/g, '_')}_${Date.now().toString(36).slice(-3)}`;
          nextCols.push({
            id: codeColId,
            label: options.codeColumnLabel,
            position: options.codeColumnPosition || 'after_no',
          });
        }
      }

      const nextItems = base.items.map((row, idx) => {
        const src = selectedItems[idx];
        if (!src) return row;
        const nextCustom = { ...(row.customFields || {}) };
        if (codeColId && src.code.trim()) {
          nextCustom[codeColId] = src.code.trim();
        }
        return {
          ...row,
          requestedDescription:
            options.overwriteDescriptions || !row.requestedDescription.trim()
              ? src.description
              : row.requestedDescription,
          offeredDescription: src.description,
          uom: src.uom || row.uom || 'PCS',
          qty: src.qty ?? row.qty,
          supplierUnitCost: src.unitPrice ?? row.supplierUnitCost,
          customFields: nextCustom,
        };
      });

      return {
        ...base,
        updatedAt: new Date().toISOString(),
        customColumns: nextCols,
        supplierQuote: {
          ...base.supplierQuote,
          supplierName: quote.supplierName || base.supplierQuote.supplierName,
          supplierQuoteRef: quote.supplierRef || base.supplierQuote.supplierQuoteRef,
          quoteDate: base.supplierQuote.quoteDate || new Date().toISOString().slice(0, 10),
          currency: quote.currency || base.supplierQuote.currency,
          incoterm: quote.incoterm || base.supplierQuote.incoterm,
          paymentTerms: quote.paymentTerms || base.supplierQuote.paymentTerms,
          deliveryLeadTime: quote.deliveryLeadTime || base.supplierQuote.deliveryLeadTime,
        },
        controls: {
          ...base.controls,
          supplierName: quote.supplierName || base.controls.supplierName,
          supplierRef: quote.supplierRef || base.controls.supplierRef,
          currency: quote.currency || base.controls.currency,
          incoterm: quote.incoterm || base.controls.incoterm,
          paymentTerms: quote.paymentTerms || base.controls.paymentTerms,
          deliveryLeadTime: quote.deliveryLeadTime || base.controls.deliveryLeadTime,
        },
        items: nextItems,
      };
    };

    if (isReadOnlyRole) {
      showToast('Switch off Safe Read-Only Mode to apply a supplier quote.');
      return;
    }

    if (options.target === 'new_job') {
      const fresh = buildUpdatedJob(
        createBlankJob(
          quote.supplierName ? `${quote.supplierName} Quote` : 'Supplier Quote Job',
          company
        )
      );
      setJobs((prev) => [fresh, ...prev]);
      setActiveJobId(fresh.id);
      queueCloudJobSync(fresh, {
        immediate: true,
        auditReason: 'Created job from supplier quote',
      });
    } else {
      mutateActiveJob(buildUpdatedJob);
    }

    setPendingSupplierQuote(null);
    setActiveTab('INTERNAL_QUOTE');
    showToast(
      `Loaded ${selectedItems.length} supplier quote items into Internal Quote & Price Calc (${quote.currency}).`
    );
  };

  const handleQuickExcelOpen = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const isPdf =
        file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
      if (isPdf) {
        const quote = await parseSupplierPdfFile(file);
        setPendingSupplierQuote(quote);
        return;
      }

      const res = await importWorkbookFromXlsx(file, company.endUsers);
      if (res.kind === 'registry') {
        handleSyncRegistry(res.entries);
        setActiveTab('REGISTRY');
        showToast(`Synced ${res.importedCount} entries from ${file.name}`);
      } else if (res.kind === 'supplier_quote') {
        setPendingSupplierQuote(res.quote);
      } else {
        handleImportJob(res.job);
        showToast(`Opened ${file.name} (${res.itemCount} items loaded)`);
      }
    } catch {
      showToast('Could not read file. Upload a supplier .pdf / .xlsx quote or Grand Group .xlsx workbook.');
    } finally {
      if (quickImportRef.current) quickImportRef.current.value = '';
    }
  };

  const handleUpdateCompany = (patch: Partial<CompanyProfile>) => {
    setCompany((prev) => {
      const next = { ...prev, ...patch };
      if (localStandaloneMode) {
        saveCompanyProfile(next);
      } else if (isCloudMode && activeOrgId) {
        updateOrgSettings(activeOrgId, { companyProfile: next }).catch((err) => {
          showToast(err instanceof Error ? err.message : 'Could not save company profile.');
        });
      }
      return next;
    });
    if (patch.endUserAddresses) {
      const updatedAddresses = patch.endUserAddresses;
      setJobs((prevJobs) =>
        prevJobs.map((job) => {
          const sel = job.controls.endUserSelect;
          const isAnyIssued = Object.values(job.issuedDocs).some((d) => d && !d.isVoided);
          if (!isAnyIssued && sel && sel !== 'OTHER' && sel in updatedAddresses) {
            const nextAddr = updatedAddresses[sel] || '';
            if (job.controls.endUserAddress !== nextAddr) {
              return {
                ...job,
                updatedAt: new Date().toISOString(),
                controls: { ...job.controls, endUserAddress: nextAddr },
              };
            }
          }
          return job;
        })
      );
    }
  };

  const handleResetRegistry = () => {
    setRegistry([]);
    saveRegistry([]);
    setJobs((prevJobs) =>
      prevJobs.map((job) => ({
        ...job,
        updatedAt: new Date().toISOString(),
        jobSequence: undefined,
        jobSequenceYear: undefined,
        issuedDocs: {},
      }))
    );
    showToast('Master Registry reset to 0. Next issued job will start at Seq #001.');
  };

  const handleResetCompany = () => {
    const defaults = createDefaultCompanyProfile();
    saveCompanyProfile(defaults);
    setCompany(defaults);
    showToast('Reset company branding, logo, and Client Directory to defaults.');
  };

  const handleUpdateThemeCustomization = (patch: Partial<ThemeCustomization>) => {
    setThemeCustomization((prev) => {
      const next = normalizeThemeCustomization({ ...prev, ...patch });
      saveThemeCustomization(next);
      return next;
    });
  };

  const handleResetThemeCustomization = () => {
    saveThemeCustomization({});
    setThemeCustomization({});
  };

  // ==================== ISSUED-DOCUMENT AMENDMENTS ====================
  const amendingJob = amendment?.jobId === activeJob.id;

  const handleStartAmendment = (prefix: DocPrefix) => {
    setAmendment({ jobId: activeJob.id, prefix, snapshot: safeDeepClone(activeJob) });
    setAmendDialog(null);
    showToast(`${prefix} unlocked for amendment. Changes are logged when you save a revision.`);
  };

  const handleDiscardAmendment = () => {
    if (!amendment) return;
    const snap = amendment.snapshot;
    setJobs((prev) => prev.map((j) => (j.id === snap.id ? snap : j)));
    setAmendment(null);
    showToast('Changes discarded. The document keeps its saved revision.');
  };

  const handleSaveRevision = (reason: string) => {
    if (!amendment) return;
    const { prefix, snapshot } = amendment;
    // Save the job that was unlocked, even if the user has since switched the active job
    const job = jobs.find((j) => j.id === amendment.jobId);
    if (!job) {
      setAmendment(null);
      setAmendDialog(null);
      return;
    }
    const changes = diffJobs(snapshot, job);
    const doc = job.issuedDocs[prefix];
    if (!doc || doc.isVoided) {
      setAmendment(null);
      setAmendDialog(null);
      return;
    }
    if (changes.length === 0) {
      setAmendment(null);
      setAmendDialog(null);
      showToast('No changes to save. The document is unchanged.');
      return;
    }

    if (isCloudMode && activeOrgId) {
      amendOrgDocument(activeOrgId, job, prefix, reason, changes)
        .then((res) => {
          const norm = normalizeJob(res.job);
          setJobs((prev) => prev.map((j) => (j.id === norm.id ? norm : j)));
          setRegistry((prev) =>
            prev.map((e) =>
              e.docNumber === res.docNumber
                ? { ...e, amount: res.amount, revision: res.revision }
                : e
            )
          );
          showToast(
            `Saved ${res.docNumber} as Rev ${res.revision}. ${changes.length} change(s) logged to organization audit trail.`
          );
          setAmendment(null);
          setAmendDialog(null);
        })
        .catch((err) => {
          showToast(err instanceof Error ? err.message : 'Could not save revision.');
        });
      return;
    }

    const now = new Date().toISOString();
    const previousRev = doc.revision ?? 1;
    const newRev = previousRev + 1;
    const newCalc = calculateJob(job);
    const docAmount =
      prefix === 'PO' ? newCalc.totalSupplierCost : prefix === 'PL' || prefix === 'DN' ? 0 : newCalc.grandTotal;

    const superseded: IssuedRevision = {
      revision: previousRev,
      docNumber: doc.docNumber,
      supersededAt: now,
      reason,
      frozenSubtotal: doc.frozenSubtotal,
      frozenVat: doc.frozenVat,
      frozenFreight: doc.frozenFreight,
      frozenGrandTotal: doc.frozenGrandTotal,
    };
    const nextDoc: IssuedDocRecord = {
      ...doc,
      revision: newRev,
      frozenSubtotal: newCalc.goodsSubtotal,
      frozenVat: newCalc.vatAmount,
      frozenFreight: newCalc.customerFreight,
      frozenGrandTotal: newCalc.grandTotal,
      history: [...(doc.history ?? []), superseded],
    };

    setJobs((prev) =>
      prev.map((j) =>
        j.id === job.id ? { ...job, updatedAt: now, issuedDocs: { ...job.issuedDocs, [prefix]: nextDoc } } : j
      )
    );
    setRegistry((prev) =>
      prev.map((e) => (e.docNumber === doc.docNumber ? { ...e, amount: docAmount, revision: newRev } : e))
    );

    const entries: AuditEntry[] = changes.map((c, i) => ({
      id: `aud-${Date.now().toString(36)}-${i}`,
      at: now,
      jobId: job.id,
      jobName: job.jobName,
      docNumber: doc.docNumber,
      prefix,
      revision: newRev,
      reason,
      field: c.field,
      from: c.from,
      to: c.to,
    }));
    if (!appendAuditLog(entries)) {
      showToast('Revision saved, but the change log could not be written (browser storage is full).');
    } else {
      showToast(`Saved ${doc.docNumber} as Rev ${newRev}. ${changes.length} change(s) logged.`);
    }
    setAmendment(null);
    setAmendDialog(null);
  };

  const handleExportAuditLog = async (format: 'json' | 'csv' | 'xlsx') => {
    let entries: AuditEntry[] = [];
    if (isCloudMode && activeOrgId) {
      try {
        const res = await fetchOrgAuditLog(activeOrgId);
        entries = res.entries;
      } catch {
        entries = loadAuditLog();
      }
    } else {
      entries = loadAuditLog();
    }
    const day = new Date().toISOString().slice(0, 10);
    if (format === 'xlsx') {
      await exportAuditLogXlsx(entries, company);
    } else if (format === 'csv') {
      const blob = new Blob([auditLogToCsv(entries)], { type: 'text/csv;charset=utf-8' });
      await shareOrDownloadBlob(blob, `GG_Audit_Log_${day}.csv`, 'text/csv');
    } else {
      const payload = { exportedAt: new Date().toISOString(), entries };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
      await shareOrDownloadBlob(blob, `GG_Audit_Log_${day}.json`, 'application/json');
    }
    showToast(`Exported ${entries.length} audit log entries (${format.toUpperCase()}).`);
  };

  const handleRestoreBackup = (
    restoredJobs: ProcurementJob[],
    restoredRegistry: RegistryEntry[],
    restoredCompany?: CompanyProfile,
    restoredThemeId?: ThemeVariantId,
    restoredThemeCustom?: ThemeCustomization,
    restoredDashPrefs?: DashboardPreferences
  ) => {
    if (restoredJobs.length > 0) {
      const normalized = restoredJobs.map((j) => normalizeJob(j));
      setJobs(normalized);
      setActiveJobId(normalized[0].id);
    }
    const cleanReg = normalizeRegistryEntries(restoredRegistry);
    setRegistry((prev) => mergeRegistryEntries(prev, cleanReg));
    if (restoredCompany) {
      const normCompany = normalizeCompanyProfile(restoredCompany);
      saveCompanyProfile(normCompany);
      setCompany(normCompany);
    }
    if (restoredThemeId) {
      setThemeId(restoredThemeId);
    }
    if (restoredThemeCustom) {
      const normCustom = normalizeThemeCustomization(restoredThemeCustom);
      saveThemeCustomization(normCustom);
      setThemeCustomization(normCustom);
    }
    if (restoredDashPrefs) {
      saveDashboardPrefs(restoredDashPrefs);
      setDashboardPrefs(restoredDashPrefs);
    }
  };

  const navItems: {
    group: string;
    items: { id: WorkspaceTab; label: string; icon: React.ReactNode; docPrefix?: DocPrefix }[];
  }[] = [
    {
      group: 'Overview',
      items: [
        ...(isCloudMode
          ? [
              {
                id: 'EXECUTIVE' as WorkspaceTab,
                label: 'Executive Live',
                icon: <Radio className="w-4 h-4" />,
              },
            ]
          : []),
        { id: 'TRACKER', label: 'Dashboard', icon: <LayoutDashboard className="w-4 h-4" /> },
      ],
    },
    {
      group: 'Internal Costing',
      items: [
        { id: 'SUPPLIER_RFQ', label: 'Supplier RFQ', icon: <Send className="w-4 h-4" /> },
        { id: 'INTERNAL_QUOTE', label: 'Internal Quote', icon: <Building2 className="w-4 h-4" /> },
        { id: 'PRICE_CALC', label: 'Price Calc (Internal)', icon: <Calculator className="w-4 h-4" /> },
      ],
    },
    {
      group: 'Trade Documents',
      items: [
        { id: 'QT', label: 'Quotation (Master)', icon: <FileText className="w-4 h-4" />, docPrefix: 'QT' },
        { id: 'PI', label: 'Proforma Invoice', icon: <FileText className="w-4 h-4" />, docPrefix: 'PI' },
        { id: 'PO', label: 'Supplier PO', icon: <ShoppingCart className="w-4 h-4" />, docPrefix: 'PO' },
        { id: 'CI', label: 'Commercial Invoice', icon: <Receipt className="w-4 h-4" />, docPrefix: 'CI' },
        { id: 'PL', label: 'Packing List', icon: <Package className="w-4 h-4" />, docPrefix: 'PL' },
        { id: 'DN', label: 'Delivery Note', icon: <Truck className="w-4 h-4" />, docPrefix: 'DN' },
      ],
    },
    {
      group: 'Ledger & Administration',
      items: [
        { id: 'REGISTRY', label: 'Master Registry', icon: <Hash className="w-4 h-4" /> },
        { id: 'FILES', label: 'Import & Export', icon: <ArrowLeftRight className="w-4 h-4" /> },
        { id: 'STUDIO', label: 'Clients & Settings', icon: <Settings className="w-4 h-4" /> },
        ...(isCloudMode
          ? [
              {
                id: 'ORG_ADMIN' as WorkspaceTab,
                label: 'Organization & Seats',
                icon: <Users className="w-4 h-4" />,
              },
            ]
          : []),
        ...(authUser?.isVendorAdmin
          ? [
              {
                id: 'VENDOR_LICENSES' as WorkspaceTab,
                label: 'Vendor License Keys',
                icon: <KeyRound className="w-4 h-4" />,
              },
            ]
          : []),
        { id: 'HELP', label: 'Help & Troubleshooting', icon: <CircleHelp className="w-4 h-4" /> },
      ],
    },
  ];

  const isDocTab = (tab: WorkspaceTab): tab is DocPrefix =>
    ['QT', 'PI', 'PO', 'PL', 'DN', 'CI'].includes(tab);

  const endUser = getResolvedEndUser(activeJob) || 'Unassigned End User';

  // Gate 1: Initial Auth check loading
  if (!authReady) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6">
        <div className="bg-white border border-slate-200 rounded-xl px-8 py-6 text-center space-y-2">
          <div className="font-display text-sm font-semibold text-slate-900">
            Connecting to Grand Group Procurement Suite...
          </div>
          <div className="text-xs text-slate-500">Verifying workspace session</div>
        </div>
      </div>
    );
  }

  // Gate 2: Standalone Vendor Console opened from Onboarding
  if (!localStandaloneMode && authUser?.isVendorAdmin && !activeOrg && onboardingVendorOpen) {
    return (
      <div className="min-h-screen bg-slate-50 p-6 max-w-6xl mx-auto">
        <VendorLicenseConsoleView
          onNotify={showToast}
          onBack={() => setOnboardingVendorOpen(false)}
        />
        {toast && (
          <div className="fixed bottom-4 left-1/2 -translate-x-1/2 z-50 rounded-lg bg-slate-900 px-4 py-2.5 text-xs font-medium text-white shadow-lg">
            {toast}
          </div>
        )}
      </div>
    );
  }

  // Gate 3: Sign-In or Organization Onboarding (unless user chose Standalone Offline Local Mode)
  if (!localStandaloneMode && (!authUser || !activeOrg)) {
    return (
      <AuthGateView
        mode={authUser ? 'org_onboarding' : 'sign_in'}
        user={authUser}
        firebaseUserEmail={firebaseUser?.email ?? null}
        pendingInvites={pendingInvites}
        authLoading={authLoading}
        authError={authError}
        onGoogleSignIn={handleGoogleSignIn}
        onEmailSignIn={handleEmailSignIn}
        onEmailSignUp={handleEmailSignUp}
        onRetrySession={handleRetrySession}
        onSignOut={handleSignOut}
        onEnterLocalOfflineMode={() => {
          setLocalStandaloneMode(true);
          setActiveTab('TRACKER');
        }}
        onOrganizationReady={async (org) => {
          if (!org || typeof org.id !== 'string') return;
          setOrganizations((prev) => [
            org,
            ...(Array.isArray(prev) ? prev : []).filter((o) => o && o.id !== org.id),
          ]);
          setActiveOrgId(org.id);
          await loadOrganizationWorkspace(org);
        }}
        onOpenVendorConsole={
          authUser?.isVendorAdmin ? () => setOnboardingVendorOpen(true) : undefined
        }
      />
    );
  }

  return (
    <div className="min-h-screen flex bg-slate-50 text-slate-900">
      {/* Mobile / iPad Sidebar Backdrop */}
      {sidebarOpen && (
        <div
          className="no-print fixed inset-0 z-40 bg-slate-900/40 lg:hidden"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      {/* Persistent Left Sidebar (260px Desktop, Slide-Over on iPad/Mobile) */}
      <aside
        className={`no-print fixed lg:sticky top-0 left-0 z-50 h-screen w-[264px] shrink-0 border-r border-slate-200 bg-white flex flex-col transition-transform duration-200 ${
          sidebarOpen ? 'translate-x-0' : '-translate-x-full lg:translate-x-0'
        }`}
      >
        {/* Brand Wordmark & Logo */}
        <div className="h-16 px-4 flex items-center justify-between border-b border-slate-200">
          <div className="flex items-center gap-2.5 min-w-0">
            <div className="w-9 h-9 rounded-md border border-slate-200 bg-white p-1 flex items-center justify-center shrink-0">
              <img
                src={company.logoDataUrl}
                alt={company.name}
                className="max-w-full max-h-full object-contain"
              />
            </div>
            <span className="font-display text-sm font-bold tracking-tight text-slate-900 truncate">
              {company.name}
            </span>
          </div>
          <button
            type="button"
            onClick={() => setSidebarOpen(false)}
            className="lg:hidden p-1.5 rounded-md text-slate-500 hover:bg-slate-100 cursor-pointer"
            aria-label="Close navigation"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Active Job Switcher */}
        <div className="p-4 border-b border-slate-200 space-y-2.5 bg-slate-50/60">
          <div className="flex items-center justify-between">
            <label htmlFor="job-select" className="text-xs font-medium text-slate-500">
              Active Job ({jobs.length})
            </label>
            {jobs.length > 1 && (
              <button
                type="button"
                onClick={handleDeleteCurrentJob}
                title="Delete current job"
                className="text-slate-400 hover:text-red-600 transition-colors cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
          <select
            id="job-select"
            value={activeJob.id}
            onChange={(e) => setActiveJobId(e.target.value)}
            className="w-full min-h-[40px] rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs font-medium text-slate-900 focus:border-slate-900 focus:outline-none"
          >
            {jobs.map((j) => (
              <option key={j.id} value={j.id}>
                {j.jobName}
              </option>
            ))}
          </select>
          <button
            type="button"
            onClick={handleCreateNewJob}
            className="w-full flex items-center justify-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-700 hover:bg-slate-100 hover:text-slate-900 transition-colors cursor-pointer min-h-[38px] whitespace-nowrap"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>New Blank Job</span>
          </button>
        </div>

        {/* Workflow Navigation */}
        <nav className="flex-1 overflow-y-auto px-3 py-4 space-y-5">
          {navItems.map((section) => (
            <div key={section.group}>
              <div className="px-2.5 mb-1.5 text-[11px] font-semibold text-slate-400">
                {section.group}
              </div>
              <div className="space-y-0.5">
                {section.items.map((item) => {
                  const isActive = activeTab === item.id;
                  const issuedInfo = item.docPrefix ? activeJob.issuedDocs[item.docPrefix] : undefined;
                  const isIssued = Boolean(issuedInfo && !issuedInfo.isVoided);

                  return (
                    <button
                      key={item.id}
                      type="button"
                      onClick={() => {
                        if (item.id === 'HELP') return openHelp();
                        setActiveTab(item.id);
                        setSidebarOpen(false);
                      }}
                      className={`w-full flex items-center justify-between gap-2 px-2.5 py-2 rounded-lg text-xs font-medium transition-colors cursor-pointer min-h-[40px] whitespace-nowrap ${
                        isActive
                          ? 'text-white'
                          : 'text-slate-700 hover:bg-slate-100 hover:text-slate-900'
                      }`}
                      style={isActive ? { backgroundColor: theme.primary } : undefined}
                    >
                      <span className="flex items-center gap-2.5 min-w-0">
                        <span className={isActive ? 'text-white' : 'text-slate-400'}>{item.icon}</span>
                        <span className="truncate">{item.label}</span>
                      </span>

                      {isIssued && (
                        <span
                          className={`font-mono-tabular text-[11px] shrink-0 ${
                            isActive ? 'text-white/90' : 'text-emerald-700 font-semibold'
                          }`}
                        >
                          #{String(issuedInfo?.sequence).padStart(3, '0')}
                        </span>
                      )}
                      {item.id === 'REGISTRY' && (
                        <span
                          className={`font-mono-tabular text-[11px] shrink-0 ${
                            isActive ? 'text-white/90' : 'text-slate-400'
                          }`}
                        >
                          #{String(nextSeq).padStart(3, '0')}
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
        </nav>

        {/* Organization Account & Role Footer */}
        <div className="p-3 border-t border-slate-200 bg-slate-50/70 space-y-2">
          {isCloudMode && activeOrg && authUser ? (
            <>
              {organizations.length > 1 && (
                <select
                  aria-label="Switch organization"
                  value={activeOrg.id}
                  onChange={(e) => {
                    const nextOrg = organizations.find((o) => o.id === e.target.value);
                    if (nextOrg) {
                      setActiveOrgId(nextOrg.id);
                      void loadOrganizationWorkspace(nextOrg);
                    }
                  }}
                  className="w-full rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs font-medium text-slate-800"
                >
                  {organizations.map((o) => (
                    <option key={o.id} value={o.id}>
                      {o.name} ({ROLE_LABELS[o.role]})
                    </option>
                  ))}
                </select>
              )}
              <div className="flex items-center justify-between gap-2 px-1">
                <div className="min-w-0">
                  <div className="text-xs font-semibold text-slate-900 truncate">
                    {authUser.displayName || authUser.email}
                  </div>
                  <div className="text-[11px] text-slate-500 truncate">
                    {ROLE_LABELS[activeOrg.role]} · {activeOrg.activeSeatsUsed}/{activeOrg.seatLimit} seats
                  </div>
                </div>
                <button
                  type="button"
                  onClick={handleSignOut}
                  title="Sign out"
                  className="p-1.5 rounded-md text-slate-500 hover:text-slate-900 hover:bg-slate-200/70 cursor-pointer shrink-0"
                >
                  <LogOut className="w-3.5 h-3.5" />
                </button>
              </div>
            </>
          ) : (
            <button
              type="button"
              onClick={() => setLocalStandaloneMode(false)}
              className="w-full inline-flex items-center justify-center gap-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 px-3 py-2 text-xs font-semibold text-slate-800 cursor-pointer"
            >
              <LogIn className="w-3.5 h-3.5" />
              <span>Sign In to Cloud Organization</span>
            </button>
          )}
        </div>
      </aside>

      {/* Main Content Shell */}
      <div className="flex-1 flex flex-col min-w-0">
        {/* Clean Single-Row Top Bar Contract (Breadcrumb + 1 Primary Action) */}
        <header
          className="no-print sticky top-0 z-30 h-16 border-b border-slate-200 backdrop-blur px-4 sm:px-6 flex items-center justify-between gap-6"
          style={{ backgroundColor: pageBarColor(activeTab, theme), color: theme.primaryText }}
        >
          <div className="flex items-center gap-3 min-w-0">
            <button
              type="button"
              onClick={() => setSidebarOpen(true)}
              className="lg:hidden p-2 -ml-1 rounded-lg text-slate-600 hover:bg-slate-100 cursor-pointer min-h-[40px] min-w-[40px] flex items-center justify-center"
              aria-label="Open menu"
            >
              <Menu className="w-5 h-5" />
            </button>
            <div className="flex items-center gap-2 text-sm min-w-0">
              <span className="hidden sm:inline opacity-75 truncate max-w-[200px]">
                {activeJob.jobName}
              </span>
              <span className="hidden sm:inline opacity-60">/</span>
              <h1 className="font-display font-semibold truncate">
                {TAB_LABELS[activeTab]}
              </h1>
            </div>
          </div>

          <div className="flex items-center gap-3 shrink-0">
            <div className="hidden md:flex items-center gap-1">
              <button
                type="button"
                onClick={handleUndo}
                disabled={historyFor(currentUndoKey()).past.length === 0}
                title={`Undo (Ctrl+Z) · ${historyFor(currentUndoKey()).past.length} of ${UNDO_LIMIT} steps · this page`}
                aria-label="Undo"
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-2.5 py-2 text-xs font-semibold text-slate-900 hover:bg-slate-50 disabled:opacity-40 cursor-pointer min-h-[40px]"
              >
                <Undo2 className="w-3.5 h-3.5" />
                <span>Undo</span>
              </button>
              <button
                type="button"
                onClick={handleRedo}
                disabled={historyFor(currentUndoKey()).future.length === 0}
                title="Redo (Ctrl+Y)"
                aria-label="Redo"
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-2.5 py-2 text-xs font-semibold text-slate-900 hover:bg-slate-50 disabled:opacity-40 cursor-pointer min-h-[40px]"
              >
                <Redo2 className="w-3.5 h-3.5" />
                <span>Redo</span>
              </button>
            </div>
            <input
              ref={quickImportRef}
              type="file"
              accept=".xlsx,.pdf,application/pdf"
              onChange={handleQuickExcelOpen}
              className="hidden"
            />
            <button
              type="button"
              onClick={openHelp}
              className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-3.5 py-2 text-xs font-semibold text-slate-900 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0"
            >
              <CircleHelp className="w-3.5 h-3.5" />
              <span>Help</span>
            </button>
            <button
              type="button"
              onClick={() => quickImportRef.current?.click()}
              className="inline-flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold transition-opacity hover:opacity-90 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0"
              style={{ backgroundColor: "#FFFFFF", color: theme.primary }}
            >
              <FolderOpen className="w-3.5 h-3.5" />
              <span>Open .xlsx / .pdf</span>
            </button>
          </div>
        </header>

        {/* Workspace Toolbar: Unboxed Metadata Summary + 3-Variant Excel Color Switcher */}
        <div className="no-print border-b border-slate-200 bg-slate-50 px-4 sm:px-6 py-2.5 flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2 text-xs text-slate-600">
            <span className="font-medium text-slate-900">{endUser}</span>
            <span aria-hidden="true" className="text-slate-300">·</span>
            <span>
              {activeJob.controls.incoterm} {activeJob.controls.destination}
            </span>
            <span aria-hidden="true" className="text-slate-300">·</span>
            <span className="font-mono-tabular">
              Markup {formatPct(activeJob.freight.selectedMarkupPct)}
            </span>
            <span aria-hidden="true" className="text-slate-300">·</span>
            <span className="font-mono-tabular font-semibold text-slate-900">
              Grand Total {formatMoney(calc.grandTotal, activeJob.controls.currency)}
            </span>
          </div>

          {/* Executive Safe Mode / Live Shortcut & Platform Version Switcher */}
          <div className="flex flex-wrap items-center gap-2.5">
            {isCloudMode && activeOrg && activeTab !== 'EXECUTIVE' && (
              <button
                type="button"
                onClick={() => setActiveTab('EXECUTIVE')}
                className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md border border-slate-200 bg-white hover:bg-slate-100 text-xs font-semibold text-slate-800 cursor-pointer whitespace-nowrap min-h-[30px]"
              >
                <span
                  className={`w-2 h-2 rounded-full ${
                    executiveStream.streamState === 'live' ? 'bg-emerald-500' : 'bg-amber-500'
                  }`}
                  aria-hidden="true"
                />
                <span>Executive Live</span>
              </button>
            )}

            {isCloudMode &&
              activeOrg &&
              (activeOrg.role === 'CEO_EXECUTIVE' || activeOrg.role === 'ORG_ADMIN') &&
              activeTab !== 'EXECUTIVE' && (
                <button
                  type="button"
                  onClick={() => setExecutiveSafeMode((v) => !v)}
                  className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md border text-xs font-semibold cursor-pointer whitespace-nowrap min-h-[30px] ${
                    executiveSafeMode
                      ? 'border-amber-300 bg-amber-50 text-amber-900'
                      : 'border-emerald-300 bg-emerald-50 text-emerald-900'
                  }`}
                >
                  {executiveSafeMode ? (
                    <>
                      <Lock className="w-3 h-3" />
                      <span>Safe Browse</span>
                    </>
                  ) : (
                    <>
                      <Unlock className="w-3 h-3" />
                      <span>Edit Mode</span>
                    </>
                  )}
                </button>
              )}

            {pendingOutboxCount > 0 && isCloudMode && (
              <button
                type="button"
                onClick={handleFlushOfflineOutbox}
                className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md border border-amber-300 bg-amber-50 text-amber-900 text-xs font-semibold cursor-pointer whitespace-nowrap min-h-[30px]"
              >
                <span>{pendingOutboxCount} Offline Queued · Sync</span>
              </button>
            )}

            <div className="inline-flex items-center p-0.5 rounded-lg bg-slate-200/80">
              <button
                type="button"
                onClick={() => setPlatformMode('windows')}
                className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap min-h-[30px] ${
                  platformMode === 'windows'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Monitor className="w-3.5 h-3.5" />
                <span>Windows</span>
              </button>
              <button
                type="button"
                onClick={() => setPlatformMode('ipad')}
                className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap min-h-[30px] ${
                  platformMode === 'ipad'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Tablet className="w-3.5 h-3.5" />
                <span>iPad / iOS</span>
              </button>
            </div>
          </div>
        </div>

        {/* Expired License or Optimistic Concurrency Conflict Banners */}
        {isCloudMode && activeOrg?.licenseExpired && (
          <div className="no-print border-b border-red-200 bg-red-50 px-6 py-2.5 flex flex-wrap items-center justify-between gap-3 text-xs text-red-900">
            <div className="flex items-center gap-2 font-medium">
              <ShieldAlert className="w-4 h-4 shrink-0" />
              <span>
                Your organization license has expired. Editing is paused until an Organization Admin
                applies a renewed signed license key.
              </span>
            </div>
            {activeOrg.role === 'ORG_ADMIN' && (
              <button
                type="button"
                onClick={() => setActiveTab('ORG_ADMIN')}
                className="px-3 py-1 rounded-md bg-red-700 hover:bg-red-800 text-white font-semibold cursor-pointer"
              >
                Apply License Key
              </button>
            )}
          </div>
        )}

        {syncConflict && isCloudMode && activeOrgId && (
          <div className="no-print border-b border-amber-300 bg-amber-50 px-6 py-2.5 flex flex-wrap items-center justify-between gap-3 text-xs text-amber-950">
            <span>{syncConflict.message}</span>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => {
                  setJobs((prev) =>
                    prev.map((j) => (j.id === syncConflict.jobId ? syncConflict.serverJob : j))
                  );
                  setSyncConflict(null);
                  showToast('Loaded latest server version of the job.');
                }}
                className="px-3 py-1 rounded-md border border-amber-400 bg-white hover:bg-amber-100 font-semibold cursor-pointer"
              >
                Load Server Latest
              </button>
              <button
                type="button"
                onClick={() => {
                  queueCloudJobSync(syncConflict.localJob, {
                    immediate: true,
                    forceOverwrite: true,
                    auditReason: 'Resolved conflict keeping local draft',
                  });
                  setSyncConflict(null);
                  showToast('Saved your local draft to the cloud workspace.');
                }}
                className="px-3 py-1 rounded-md bg-amber-700 hover:bg-amber-800 text-white font-semibold cursor-pointer"
              >
                Keep My Draft
              </button>
            </div>
          </div>
        )}

        {/* iPad Quick Document Selector Strip when viewing Trade Documents */}
        {platformMode === 'ipad' && isDocTab(activeTab) && (
          <div className="no-print border-b border-slate-200 bg-white px-4 py-2 overflow-x-auto">
            <div className="flex items-center gap-1.5 min-w-max">
              {(['QT', 'PI', 'PO', 'CI', 'PL', 'DN'] as DocPrefix[]).map((doc) => {
                const isSelected = activeTab === doc;
                const issuedInfo = activeJob.issuedDocs[doc];
                return (
                  <button
                    key={doc}
                    type="button"
                    onClick={() => setActiveTab(doc)}
                    className={`px-3.5 py-2 rounded-lg text-xs font-semibold transition-colors cursor-pointer min-h-[40px] whitespace-nowrap ${
                      isSelected
                        ? 'text-white'
                        : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                    }`}
                    style={isSelected ? { backgroundColor: theme.primary } : undefined}
                  >
                    {DOC_META[doc].shortName}
                    {issuedInfo && !issuedInfo.isVoided ? ` · #${String(issuedInfo.sequence).padStart(3, '0')}` : ''}
                  </button>
                );
              })}
            </div>
          </div>
        )}

        {/* Active Viewport Content */}
        <main className={`flex-1 p-4 sm:p-6 max-w-[1440px] w-full mx-auto ${platformMode === 'ipad' ? 'pb-24' : ''}`}>
          {workspaceLoading && (
            <div className="mb-4 rounded-lg border border-slate-200 bg-white px-4 py-2.5 text-xs text-slate-600">
              Synchronizing organization workspace...
            </div>
          )}

          {activeTab === 'EXECUTIVE' && isCloudMode && activeOrg && (
            <ExecutiveLiveView
              snapshot={executiveStream.snapshot}
              loading={executiveStream.loading}
              error={executiveStream.error}
              streamState={executiveStream.streamState}
              lastPulseAt={executiveStream.lastPulseAt}
              role={activeOrg.role}
              executiveSafeMode={executiveSafeMode}
              onToggleSafeMode={() => setExecutiveSafeMode((v) => !v)}
              onRefreshSnapshot={executiveStream.refreshSnapshot}
              onOpenJobInFullApp={(jobId, tab) => {
                handleSelectJobAndTab(jobId, tab);
              }}
              onOpenOrgSeats={() => setActiveTab('ORG_ADMIN')}
              onNotify={showToast}
            />
          )}

          {activeTab === 'ORG_ADMIN' && isCloudMode && activeOrg && (
            <OrgAdminView
              organization={activeOrg}
              offlineLease={offlineLease}
              pendingOutboxCount={pendingOutboxCount}
              onSyncNow={handleFlushOfflineOutbox}
              onWorkspaceReloaded={async () => {
                await bootstrapSession();
              }}
              onNotify={showToast}
            />
          )}

          {activeTab === 'VENDOR_LICENSES' && authUser?.isVendorAdmin && (
            <VendorLicenseConsoleView onNotify={showToast} />
          )}

          {activeTab === 'TRACKER' && (
            <JobTrackerDashboard
              jobs={jobs}
              activeJobId={activeJob.id}
              registry={registry}
              theme={theme}
              company={company}
              prefs={dashboardPrefs}
              onUpdatePrefs={handleUpdateDashboardPrefs}
              onSelectJobAndTab={handleSelectJobAndTab}
              onCreateNewJob={handleCreateNewJob}
              onNotify={showToast}
            />
          )}

          {activeTab === 'SUPPLIER_RFQ' && (
            <SupplierRfqView
              job={activeJob}
              company={company}
              theme={theme}
              onUpdateRfq={handleUpdateRfq}
              onFillEmptyRows={handleFillEmptyRows}
              onNotify={showToast}
            />
          )}

          {activeTab === 'INTERNAL_QUOTE' && (
            <InternalQuoteView
              job={activeJob}
              amending={amendingJob}
              calc={calc}
              theme={theme}
              onUpdateSupplierQuote={handleUpdateSupplierQuote}
              onUpdateControls={handleUpdateControls}
              onUpdateItem={handleUpdateItem}
              onOpenImportFile={() => quickImportRef.current?.click()}
              onNavigateTab={setActiveTab}
              onRequestConfirm={setConfirmDialog}
              onNotify={showToast}
            />
          )}

          {isDocTab(activeTab) && (
            <DocumentWorkspaceView
              key={`${activeJob.id}-${activeTab}`}
              job={activeJob}
              prefix={activeTab}
              calc={calc}
              theme={theme}
              company={company}
              platformMode={platformMode}
              nextNewJobSeq={nextSeq}
              onUpdateJobName={handleUpdateJobName}
              onUpdateInquiryType={handleUpdateInquiryType}
              onUpdateControls={handleUpdateControls}
              onUpdateWorkflow={handleUpdateWorkflow}
              onUpdateCustomColumns={handleUpdateCustomColumns}
              onUpdateColumnOrder={handleUpdateColumnOrder}
              onUpdateItem={handleUpdateItem}
              onIssueDocument={handleIssueDocument}
              amending={amendingJob}
              amendingPrefix={amendingJob ? amendment?.prefix ?? null : null}
              onRequestUnlock={(p) => setAmendDialog({ mode: "unlock", prefix: p })}
              onRequestSave={(p) => setAmendDialog({ mode: "save", prefix: p })}
              onDiscardAmendment={handleDiscardAmendment}
              onDuplicateAsNewJob={handleDuplicateAsNewJob}
              onNavigateTab={setActiveTab}
              onRequestConfirm={setConfirmDialog}
              onNotify={showToast}
            />
          )}

          {activeTab === 'PRICE_CALC' && (
            <PriceCalcView
              job={activeJob}
              amending={amendingJob}
              calc={calc}
              theme={theme}
              platformMode={platformMode}
              onUpdateItem={handleUpdateItem}
              onUpdateFreight={handleUpdateFreight}
              onUpdateControls={handleUpdateControls}
              onNavigateTab={setActiveTab}
              onRequestConfirm={setConfirmDialog}
            />
          )}

          {activeTab === 'REGISTRY' && (
            <RegistryView
              entries={registry}
              nextSeq={nextSeq}
              theme={theme}
              company={company}
              onVoidEntry={handleVoidRegistryEntry}
              onSyncRegistryEntries={handleSyncRegistry}
              onResetRegistry={handleResetRegistry}
              onExportAuditLog={handleExportAuditLog}
              onRequestConfirm={setConfirmDialog}
              onNotify={showToast}
            />
          )}

          {activeTab === 'HELP' && (
            <HelpView
              key={`${helpTopicId}-${helpReturnTab}`}
              initialTopicId={helpTopicId}
              returnLabel={TAB_LABELS[helpReturnTab]}
              onBack={() => setActiveTab(helpReturnTab)}
              accentColor={theme.primary}
              accentText={theme.primaryText}
            />
          )}

          {activeTab === 'FILES' && (
            <FilesHubView
              activeJob={activeJob}
              allJobs={jobs}
              registry={registry}
              themeId={themeId}
              theme={theme}
              themeCustomization={themeCustomization}
              company={company}
              dashboardPrefs={dashboardPrefs}
              onImportJob={handleImportJob}
              onOpenSupplierQuote={setPendingSupplierQuote}
              onSyncRegistry={handleSyncRegistry}
              onRestoreBackup={handleRestoreBackup}
              onRequestConfirm={setConfirmDialog}
              onNotify={showToast}
            />
          )}

          {activeTab === 'STUDIO' && (
            <ExcelStudioView
              themeId={themeId}
              theme={theme}
              themeCustomization={themeCustomization}
              company={company}
              onSelectTheme={setThemeId}
              onUpdateThemeCustomization={handleUpdateThemeCustomization}
              onResetThemeCustomization={handleResetThemeCustomization}
              onUpdateCompany={handleUpdateCompany}
              onResetCompany={handleResetCompany}
              onRequestConfirm={setConfirmDialog}
              onNotify={showToast}
            />
          )}
        </main>
      </div>

      {/* Reusable In-App Confirmation Dialog for Destructive Actions */}
      <ConfirmDialog
        request={confirmDialog}
        theme={theme}
        onClose={() => setConfirmDialog(null)}
      />

      {/* Supplier Quote (.pdf / any .xlsx / Pasted Text) Review & Load Modal */}
      {pendingSupplierQuote && (
        <SupplierQuoteImportModal
          initialQuote={pendingSupplierQuote}
          activeJobName={activeJob.jobName}
          theme={theme}
          onApply={handleApplySupplierQuote}
          onClose={() => setPendingSupplierQuote(null)}
        />
      )}

      {/* iPadOS Bottom Touch Navigation Dock */}
      {platformMode === 'ipad' && (
        <nav
          aria-label="iPad Navigation"
          className="no-print fixed bottom-0 inset-x-0 z-40 border-t border-slate-200 bg-white/95 backdrop-blur px-2 pt-1.5 pb-[max(0.5rem,env(safe-area-inset-bottom))]"
        >
          <div className="max-w-2xl mx-auto grid grid-cols-5 gap-1">
            {[
              { id: 'TRACKER' as WorkspaceTab, label: 'Dashboard', icon: <LayoutDashboard className="w-5 h-5" />, active: activeTab === 'TRACKER' },
              { id: 'INTERNAL_QUOTE' as WorkspaceTab, label: 'Int. Quote', icon: <Building2 className="w-5 h-5" />, active: activeTab === 'INTERNAL_QUOTE' },
              { id: 'QT' as WorkspaceTab, label: 'Docs', icon: <Receipt className="w-5 h-5" />, active: isDocTab(activeTab) },
              { id: 'PRICE_CALC' as WorkspaceTab, label: 'Price Calc', icon: <Calculator className="w-5 h-5" />, active: activeTab === 'PRICE_CALC' },
              { id: 'REGISTRY' as WorkspaceTab, label: 'Registry', icon: <Hash className="w-5 h-5" />, active: activeTab === 'REGISTRY' },
            ].map((item) => (
              <button
                key={item.id}
                type="button"
                onClick={() => setActiveTab(item.id)}
                className={`flex flex-col items-center justify-center gap-1 rounded-lg py-1.5 min-h-[48px] text-[11px] font-medium transition-colors cursor-pointer ${
                  item.active ? 'text-white' : 'text-slate-600 hover:bg-slate-100'
                }`}
                style={item.active ? { backgroundColor: theme.primary } : undefined}
              >
                {item.icon}
                <span className="whitespace-nowrap">{item.label}</span>
              </button>
            ))}
          </div>
        </nav>
      )}

      {/* Transient Notification Toast */}
      {toast && (
        <div className={`no-print fixed ${platformMode === 'ipad' ? 'bottom-20' : 'bottom-4'} left-1/2 -translate-x-1/2 z-50 flex items-center gap-2.5 rounded-lg border border-slate-800 bg-slate-900 px-4 py-2.5 text-xs font-medium text-white shadow-lg`}>
          <FileSpreadsheet className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toast}</span>
        </div>
      )}

      {amendDialog && (
        <AmendmentDialog
          mode={amendDialog.mode}
          docNumber={activeJob.issuedDocs[amendDialog.prefix]?.docNumber ?? amendDialog.prefix}
          currentRevision={activeJob.issuedDocs[amendDialog.prefix]?.revision ?? 1}
          otherIssuedDocs={Object.values(activeJob.issuedDocs)
            .filter((d): d is IssuedDocRecord => Boolean(d && !d.isVoided && d.prefix !== amendDialog.prefix))
            .map((d) => d.prefix)}
          verifyPassword={verifyAdminPassword}
          onConfirmUnlock={() => handleStartAmendment(amendDialog.prefix)}
          onConfirmSave={handleSaveRevision}
          onCancel={() => setAmendDialog(null)}
        />
      )}
      <OfflineIndicator />
    </div>
  );
}
