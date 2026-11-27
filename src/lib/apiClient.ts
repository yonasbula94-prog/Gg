import { auth } from './firebase';
import type {
  CompanyProfile,
  DashboardPreferences,
  DocPrefix,
  FieldChange,
  ProcurementJob,
  RegistryEntry,
  ThemeCustomization,
  ThemeVariantId,
  AuditEntry,
} from '../types/procurement';
import type {
  ExecutiveSnapshot,
  LicensePlanTier,
  OrgLicenseRecord,
  OrgMemberRecord,
  OrgRole,
  OrgWorkspaceSnapshot,
  OrganizationSummary,
  SeatStatus,
  SessionBootstrapResponse,
} from '../types/enterprise';

const API_BASE = (import.meta.env.VITE_API_BASE_URL || '').replace(/\/+$/, '');

export function getApiBaseUrl(): string {
  return API_BASE;
}

export class ApiError extends Error {
  status: number;
  code?: string;
  data?: unknown;

  constructor(message: string, status: number, code?: string, data?: unknown) {
    super(message);
    this.name = 'ApiError';
    this.status = status;
    this.code = code;
    this.data = data;
  }
}

async function getAuthHeader(explicitToken?: string): Promise<Record<string, string>> {
  if (explicitToken) {
    return {
      'X-Firebase-Auth': explicitToken,
    };
  }
  const current = auth.currentUser;
  if (!current) {
    throw new ApiError('Please sign in to continue.', 401);
  }
  const idToken = await current.getIdToken();
  return {
    'X-Firebase-Auth': idToken,
  };
}

async function apiRequest<T>(
  path: string,
  init: RequestInit = {},
  options?: { explicitToken?: string }
): Promise<T> {
  const headers: Record<string, string> = {
    Accept: 'application/json',
    ...(await getAuthHeader(options?.explicitToken)),
    ...(init.body ? { 'Content-Type': 'application/json' } : {}),
    ...((init.headers as Record<string, string>) || {}),
  };

  let response: Response;
  try {
    response = await fetch(`${API_BASE}${path}`, {
      cache: 'no-store',
      credentials: 'include',
      ...init,
      headers,
    });
  } catch {
    throw new ApiError(
      'Could not connect to the workspace server. Check your internet connection and try again.',
      0,
      'NETWORK_ERROR'
    );
  }

  const contentType = response.headers.get('content-type') || 'unknown';
  const text = await response.text();
  let parsed: Record<string, unknown> | null = null;
  if (text) {
    try {
      const json = JSON.parse(text);
      if (json && typeof json === 'object') {
        parsed = json as Record<string, unknown>;
      }
    } catch {
      parsed = null;
    }
  }

  if (!response.ok) {
    throw new ApiError(
      parsed && typeof parsed.error === 'string'
        ? parsed.error
        : `Request failed (HTTP ${response.status})`,
      response.status,
      parsed && typeof parsed.code === 'string' ? parsed.code : undefined,
      parsed ?? text
    );
  }

  if (parsed === null) {
    throw new ApiError(
      response.redirected
        ? 'Your browser session was redirected before reaching the workspace API. Please reload the page and sign in again.'
        : `Workspace API returned a non-JSON response (HTTP ${response.status}, ${contentType.split(';')[0]}). Please refresh and try again.`,
      502,
      'INVALID_JSON_RESPONSE'
    );
  }

  return parsed as T;
}

export async function fetchSessionBootstrap(
  explicitToken?: string
): Promise<SessionBootstrapResponse> {
  return apiRequest<SessionBootstrapResponse>(
    '/api/auth/session',
    {},
    { explicitToken }
  );
}

export async function createOrganization(params: {
  name: string;
  licenseKey?: string;
}): Promise<{ organization: OrganizationSummary }> {
  return apiRequest('/api/orgs', {
    method: 'POST',
    body: JSON.stringify(params),
  });
}

export async function acceptInvite(
  membershipId: string
): Promise<{ organization: OrganizationSummary }> {
  return apiRequest(`/api/orgs/invites/${encodeURIComponent(membershipId)}/accept`, {
    method: 'POST',
  });
}

export async function fetchOrgWorkspace(orgId: string): Promise<OrgWorkspaceSnapshot> {
  return apiRequest<OrgWorkspaceSnapshot>(`/api/orgs/${encodeURIComponent(orgId)}/workspace`);
}

export async function updateOrgSettings(
  orgId: string,
  patch: {
    companyProfile?: CompanyProfile;
    themeVariant?: ThemeVariantId;
    themeCustomization?: ThemeCustomization;
    dashboardPrefs?: DashboardPreferences;
  }
): Promise<{ ok: boolean }> {
  return apiRequest(`/api/orgs/${encodeURIComponent(orgId)}/settings`, {
    method: 'PATCH',
    body: JSON.stringify(patch),
  });
}

export async function importLocalWorkspace(
  orgId: string,
  payload: {
    jobs: ProcurementJob[];
    registry: RegistryEntry[];
    companyProfile?: CompanyProfile;
  }
): Promise<{ ok: boolean; importedJobsCount: number; importedRegistryCount: number }> {
  return apiRequest(`/api/orgs/${encodeURIComponent(orgId)}/import-local`, {
    method: 'POST',
    body: JSON.stringify(payload),
  });
}

export async function fetchOrgMembers(orgId: string): Promise<{
  organization: Omit<OrganizationSummary, 'role' | 'seatStatus'> | null;
  members: OrgMemberRecord[];
  licenses: OrgLicenseRecord[];
}> {
  return apiRequest(`/api/orgs/${encodeURIComponent(orgId)}/members`);
}

export async function activateOrgLicense(
  orgId: string,
  licenseKey: string
): Promise<{
  ok: boolean;
  planTier: LicensePlanTier;
  seatLimit: number;
  licenseExpiresAt: string;
  licenseId: string;
}> {
  return apiRequest(`/api/orgs/${encodeURIComponent(orgId)}/license/activate`, {
    method: 'POST',
    body: JSON.stringify({ licenseKey }),
  });
}

export async function inviteOrgMember(
  orgId: string,
  params: { email: string; displayName?: string; role: OrgRole }
): Promise<{ member: OrgMemberRecord }> {
  return apiRequest(`/api/orgs/${encodeURIComponent(orgId)}/members/invite`, {
    method: 'POST',
    body: JSON.stringify(params),
  });
}

export async function updateOrgMember(
  orgId: string,
  memberId: string,
  patch: { role?: OrgRole; seatStatus?: SeatStatus }
): Promise<{ member: OrgMemberRecord }> {
  return apiRequest(
    `/api/orgs/${encodeURIComponent(orgId)}/members/${encodeURIComponent(memberId)}`,
    {
      method: 'PATCH',
      body: JSON.stringify(patch),
    }
  );
}

export async function removeOrgMember(orgId: string, memberId: string): Promise<{ ok: boolean }> {
  return apiRequest(
    `/api/orgs/${encodeURIComponent(orgId)}/members/${encodeURIComponent(memberId)}`,
    {
      method: 'DELETE',
    }
  );
}

export async function saveOrgJob(
  orgId: string,
  job: ProcurementJob,
  options?: { baseVersion?: number; forceOverwrite?: boolean; auditReason?: string }
): Promise<{ job: ProcurementJob; version: number }> {
  return apiRequest(`/api/orgs/${encodeURIComponent(orgId)}/jobs/${encodeURIComponent(job.id)}`, {
    method: 'PUT',
    body: JSON.stringify({
      job,
      baseVersion: options?.baseVersion,
      forceOverwrite: options?.forceOverwrite,
      auditReason: options?.auditReason,
    }),
  });
}

export async function deleteOrgJob(orgId: string, jobId: string): Promise<{ ok: boolean }> {
  return apiRequest(`/api/orgs/${encodeURIComponent(orgId)}/jobs/${encodeURIComponent(jobId)}`, {
    method: 'DELETE',
  });
}

export async function issueOrgDocument(
  orgId: string,
  job: ProcurementJob,
  prefix: DocPrefix
): Promise<{
  job: ProcurementJob;
  registryEntry: RegistryEntry;
  sequence: number;
  docNumber: string;
}> {
  return apiRequest(
    `/api/orgs/${encodeURIComponent(orgId)}/jobs/${encodeURIComponent(job.id)}/issue`,
    {
      method: 'POST',
      body: JSON.stringify({ job, prefix }),
    }
  );
}

export async function amendOrgDocument(
  orgId: string,
  job: ProcurementJob,
  prefix: DocPrefix,
  reason: string,
  changes: FieldChange[]
): Promise<{ job: ProcurementJob; revision: number; docNumber: string; amount: number }> {
  return apiRequest(
    `/api/orgs/${encodeURIComponent(orgId)}/jobs/${encodeURIComponent(job.id)}/amend`,
    {
      method: 'POST',
      body: JSON.stringify({ job, prefix, reason, changes }),
    }
  );
}

export async function voidOrgRegistryEntry(
  orgId: string,
  docNumber: string,
  reason: string
): Promise<{ ok: boolean; entry: RegistryEntry }> {
  return apiRequest(
    `/api/orgs/${encodeURIComponent(orgId)}/registry/${encodeURIComponent(docNumber)}/void`,
    {
      method: 'POST',
      body: JSON.stringify({ reason }),
    }
  );
}

export async function fetchOrgAuditLog(orgId: string): Promise<{ entries: AuditEntry[] }> {
  return apiRequest(`/api/orgs/${encodeURIComponent(orgId)}/audit`);
}

export async function fetchExecutiveSnapshot(orgId: string): Promise<ExecutiveSnapshot> {
  return apiRequest<ExecutiveSnapshot>(
    `/api/orgs/${encodeURIComponent(orgId)}/executive/snapshot`
  );
}

export async function requestStreamTicket(orgId: string): Promise<{
  ticket: string;
  expiresInSeconds: number;
  sseUrl: string;
  wsPath: string;
}> {
  return apiRequest(`/api/orgs/${encodeURIComponent(orgId)}/stream-ticket`, {
    method: 'POST',
  });
}

export async function fetchVendorLicenses(): Promise<{
  publicKeyPem: string;
  licenses: OrgLicenseRecord[];
}> {
  return apiRequest('/api/vendor/licenses');
}

export async function mintVendorLicense(params: {
  targetOrgName: string;
  planTier: LicensePlanTier;
  seatLimit: number;
  validDays: number;
  notes?: string;
}): Promise<{ license: OrgLicenseRecord }> {
  return apiRequest('/api/vendor/licenses', {
    method: 'POST',
    body: JSON.stringify(params),
  });
}

export async function revokeVendorLicense(
  licenseId: string
): Promise<{ ok: boolean; id: string; status: string }> {
  return apiRequest(`/api/vendor/licenses/${encodeURIComponent(licenseId)}/revoke`, {
    method: 'POST',
  });
}
