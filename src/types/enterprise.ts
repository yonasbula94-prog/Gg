import type {
  CompanyProfile,
  DashboardPreferences,
  DocPrefix,
  InquiryType,
  ProcurementJob,
  RegistryEntry,
  ThemeCustomization,
  ThemeVariantId,
  WorkspaceTab,
} from './procurement';

export type OrgRole =
  | 'ORG_ADMIN'
  | 'CEO_EXECUTIVE'
  | 'PROCUREMENT_OFFICER'
  | 'VIEWER';

export type SeatStatus = 'ACTIVE' | 'INVITED' | 'SUSPENDED';

export type LicensePlanTier = 'TRIAL' | 'STARTER' | 'BUSINESS' | 'ENTERPRISE';

export interface AuthUserProfile {
  uid: string;
  email: string;
  displayName: string;
  photoUrl: string;
  isVendorAdmin: boolean;
}

export interface OrganizationSummary {
  id: string;
  name: string;
  orgCode: string;
  planTier: LicensePlanTier;
  seatLimit: number;
  activeSeatsUsed: number;
  invitedSeatsUsed: number;
  licenseExpiresAt: string;
  licenseExpired: boolean;
  daysUntilExpiry: number;
  role: OrgRole;
  seatStatus: SeatStatus;
}

export interface PendingInvite {
  membershipId: string;
  organizationId: string;
  organizationName: string;
  orgCode: string;
  role: OrgRole;
  invitedAt: string;
}

export interface SessionBootstrapResponse {
  user: AuthUserProfile;
  organizations: OrganizationSummary[];
  pendingInvites: PendingInvite[];
}

export interface OrgMemberRecord {
  id: string;
  organizationId: string;
  userUid: string | null;
  invitedEmail: string;
  displayName: string;
  role: OrgRole;
  seatStatus: SeatStatus;
  lastActiveAt: string | null;
  createdAt: string;
}

export interface OrgLicenseRecord {
  id: string;
  organizationId: string | null;
  targetOrgName: string;
  planTier: LicensePlanTier;
  seatLimit: number;
  issuedAt: string;
  expiresAt: string;
  licenseKey: string;
  status: 'ISSUED' | 'ACTIVATED' | 'REVOKED';
  activatedAt: string | null;
  notes: string;
}

export interface OfflineLeaseToken {
  orgId: string;
  orgName: string;
  uid: string;
  email: string;
  role: OrgRole;
  seatLimit: number;
  planTier: LicensePlanTier;
  issuedAt: string;
  expiresAt: string;
  signature: string;
}

export interface OrgWorkspaceSnapshot {
  organization: OrganizationSummary;
  companyProfile: CompanyProfile;
  themeVariant: ThemeVariantId;
  themeCustomization: ThemeCustomization;
  dashboardPrefs: DashboardPreferences;
  jobs: ProcurementJob[];
  registry: RegistryEntry[];
  nextSequence: number;
  offlineLease: OfflineLeaseToken;
}

export interface ExecutiveStageMetric {
  id: string;
  label: string;
  count: number;
  pct: number;
  filterTarget?: string;
}

export interface ExecutivePerformanceBucket {
  key: string;
  label: string;
  jobsCreated: number;
  docsIssued: number;
  value: number;
  profit: number;
}

export interface ExecutiveJobRow {
  jobId: string;
  jobName: string;
  jobSequence?: number;
  inquiryType: InquiryType;
  client: string;
  supplier: string;
  currency: string;
  pipelineValue: number;
  estimatedProfit: number;
  marginPct: number;
  completionPct: number;
  currentStageLabel: string;
  nextTab: WorkspaceTab;
  nextActionLabel: string;
  isComplete: boolean;
  preparedBy: string;
  updatedAt: string;
  issuedPrefixes: DocPrefix[];
}

export interface ExecutiveActivityEvent {
  id: string;
  at: string;
  actorName: string;
  eventType: string;
  jobId: string;
  jobName: string;
  docNumber: string;
  prefix: string;
  summary: string;
}

export interface ExecutiveSnapshot {
  organizationId: string;
  organizationName: string;
  generatedAt: string;
  activeListeners: number;
  seatSummary: {
    used: number;
    limit: number;
    planTier: LicensePlanTier;
    daysUntilExpiry: number;
  };
  kpis: {
    totalJobs: number;
    activeJobs: number;
    completedJobs: number;
    rfqCount: number;
    tenderCount: number;
    rfqValue: number;
    tenderValue: number;
    totalPipelineValue: number;
    totalEstimatedProfit: number;
    avgMarginPct: number;
    avgCompletionPct: number;
    docsIssuedCount: number;
  };
  stageFunnel: ExecutiveStageMetric[];
  weeklyPerformance: ExecutivePerformanceBucket[];
  monthlyPerformance: ExecutivePerformanceBucket[];
  jobs: ExecutiveJobRow[];
  recentActivity: ExecutiveActivityEvent[];
}

export type StreamConnectionState = 'connecting' | 'live' | 'reconnecting' | 'offline';

export interface StreamEnvelope {
  type:
    | 'connected'
    | 'executive:snapshot'
    | 'job:updated'
    | 'job:deleted'
    | 'doc:issued'
    | 'registry:updated'
    | 'activity:new'
    | 'seat:updated'
    | 'seat:revoked'
    | 'heartbeat';
  orgId: string;
  at: string;
  payload: unknown;
}

export interface OfflineOutboxItem {
  id: string;
  orgId: string;
  kind: 'upsert_job' | 'delete_job' | 'update_company' | 'update_prefs';
  jobId?: string;
  payload: unknown;
  queuedAt: string;
  attempts: number;
  lastError?: string;
}

export const ROLE_LABELS: Record<OrgRole, string> = {
  ORG_ADMIN: 'Organization Admin',
  CEO_EXECUTIVE: 'CEO / Executive',
  PROCUREMENT_OFFICER: 'Procurement Officer',
  VIEWER: 'Read-Only Viewer',
};
