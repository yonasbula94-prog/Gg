import type {
  OfflineOutboxItem,
  OrgWorkspaceSnapshot,
} from '../types/enterprise';
import type { ProcurementJob } from '../types/procurement';
import { ApiError, deleteOrgJob, saveOrgJob, updateOrgSettings } from '../lib/apiClient';

const CACHE_PREFIX = 'gg_cloud_cache_v1_';
const OUTBOX_PREFIX = 'gg_cloud_outbox_v1_';

export function isTauriDesktop(): boolean {
  return typeof window !== 'undefined' && '__TAURI_INTERNALS__' in window;
}

export function saveOrgCacheSnapshot(orgId: string, snapshot: OrgWorkspaceSnapshot): void {
  try {
    localStorage.setItem(`${CACHE_PREFIX}${orgId}`, JSON.stringify(snapshot));
  } catch {
    // Ignore storage quota errors cleanly
  }
}

export function loadOrgCacheSnapshot(orgId: string): OrgWorkspaceSnapshot | null {
  try {
    const raw = localStorage.getItem(`${CACHE_PREFIX}${orgId}`);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as OrgWorkspaceSnapshot;
    if (!parsed || !parsed.organization || !parsed.offlineLease) return null;
    // Validate 7-day offline lease expiration
    const leaseExp = new Date(parsed.offlineLease.expiresAt).getTime();
    if (!Number.isFinite(leaseExp) || leaseExp <= Date.now()) {
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

export function loadOfflineOutbox(orgId: string): OfflineOutboxItem[] {
  try {
    const raw = localStorage.getItem(`${OUTBOX_PREFIX}${orgId}`);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? (parsed as OfflineOutboxItem[]) : [];
  } catch {
    return [];
  }
}

export function saveOfflineOutbox(orgId: string, items: OfflineOutboxItem[]): void {
  try {
    localStorage.setItem(`${OUTBOX_PREFIX}${orgId}`, JSON.stringify(items));
  } catch {
    // Ignore quota errors
  }
}

export function enqueueOfflineMutation(
  orgId: string,
  item: Omit<OfflineOutboxItem, 'id' | 'orgId' | 'queuedAt' | 'attempts'>
): OfflineOutboxItem[] {
  const existing = loadOfflineOutbox(orgId);
  // Coalesce repeated draft edits for the same jobId or settings kind
  const filtered = existing.filter((entry) => {
    if (item.kind === 'upsert_job' && entry.kind === 'upsert_job' && entry.jobId === item.jobId) {
      return false;
    }
    if (item.kind === 'update_company' && entry.kind === 'update_company') return false;
    if (item.kind === 'update_prefs' && entry.kind === 'update_prefs') return false;
    return true;
  });

  const next: OfflineOutboxItem[] = [
    ...filtered,
    {
      ...item,
      id: `out-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`,
      orgId,
      queuedAt: new Date().toISOString(),
      attempts: 0,
    },
  ];
  saveOfflineOutbox(orgId, next);
  return next;
}

export interface OutboxFlushResult {
  syncedCount: number;
  remaining: OfflineOutboxItem[];
  conflicts: {
    jobId: string;
    localJob: ProcurementJob;
    serverJob: ProcurementJob;
    message: string;
  }[];
}

export async function flushOfflineOutbox(orgId: string): Promise<OutboxFlushResult> {
  const queue = loadOfflineOutbox(orgId);
  if (queue.length === 0) {
    return { syncedCount: 0, remaining: [], conflicts: [] };
  }

  const remaining: OfflineOutboxItem[] = [];
  const conflicts: OutboxFlushResult['conflicts'] = [];
  let syncedCount = 0;

  for (const item of queue) {
    try {
      if (item.kind === 'upsert_job') {
        const job = item.payload as ProcurementJob;
        await saveOrgJob(orgId, job, { baseVersion: job.version });
        syncedCount += 1;
      } else if (item.kind === 'delete_job' && item.jobId) {
        await deleteOrgJob(orgId, item.jobId);
        syncedCount += 1;
      } else if (item.kind === 'update_company' || item.kind === 'update_prefs') {
        await updateOrgSettings(orgId, item.payload as Record<string, unknown>);
        syncedCount += 1;
      }
    } catch (err) {
      if (err instanceof ApiError && err.status === 409 && item.kind === 'upsert_job') {
        const data = err.data as { serverJob?: ProcurementJob } | undefined;
        if (data?.serverJob) {
          conflicts.push({
            jobId: item.jobId || (item.payload as ProcurementJob).id,
            localJob: item.payload as ProcurementJob,
            serverJob: data.serverJob,
            message: err.message,
          });
          continue;
        }
      }
      remaining.push({
        ...item,
        attempts: item.attempts + 1,
        lastError: err instanceof Error ? err.message : 'Sync failed',
      });
    }
  }

  saveOfflineOutbox(orgId, remaining);
  return { syncedCount, remaining, conflicts };
}
