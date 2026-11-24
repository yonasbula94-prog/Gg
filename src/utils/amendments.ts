import { FieldChange, JobItemRow, ProcurementJob } from '../types/procurement';

const ITEM_FIELDS: { key: keyof JobItemRow; label: string }[] = [
  { key: 'requestedDescription', label: 'Requested description' },
  { key: 'offeredDescription', label: 'Offered description' },
  { key: 'uom', label: 'UoM' },
  { key: 'qty', label: 'Qty' },
  { key: 'supplierUnitCost', label: 'Supplier unit cost' },
  { key: 'boxes', label: 'Boxes' },
  { key: 'boxLengthCm', label: 'Box length (cm)' },
  { key: 'boxWidthCm', label: 'Box width (cm)' },
  { key: 'boxHeightCm', label: 'Box height (cm)' },
  { key: 'kgPerBox', label: 'Kg per box' },
  { key: 'cartonNo', label: 'Carton no.' },
  { key: 'netWtKg', label: 'Net weight (kg)' },
  { key: 'grossWtKg', label: 'Gross weight (kg)' },
];

function str(v: unknown): string {
  if (v === null || v === undefined || v === '') return '—';
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
}

/** Field-level differences between two versions of the same job. */
export function diffJobs(before: ProcurementJob, after: ProcurementJob): FieldChange[] {
  const changes: FieldChange[] = [];

  const ctrlKeys = Object.keys(after.controls) as (keyof ProcurementJob['controls'])[];
  for (const key of ctrlKeys) {
    if (str(before.controls[key]) !== str(after.controls[key])) {
      changes.push({ field: `Control · ${String(key)}`, from: str(before.controls[key]), to: str(after.controls[key]) });
    }
  }

  const freightKeys = Object.keys(after.freight) as (keyof ProcurementJob['freight'])[];
  for (const key of freightKeys) {
    if (str(before.freight[key]) !== str(after.freight[key])) {
      changes.push({ field: `Freight · ${String(key)}`, from: str(before.freight[key]), to: str(after.freight[key]) });
    }
  }

  if (str(before.customColumns) !== str(after.customColumns)) {
    changes.push({ field: 'Custom columns', from: str(before.customColumns), to: str(after.customColumns) });
  }

  after.items.forEach((row, idx) => {
    const prev = before.items[idx];
    if (!prev) return;
    for (const { key, label } of ITEM_FIELDS) {
      if (str(prev[key]) !== str(row[key])) {
        changes.push({
          field: `Row ${row.rowNumber} · ${label}`,
          from: str(prev[key]),
          to: str(row[key]),
        });
      }
    }
  });

  return changes;
}

/** A job is locked once any of its documents has been issued and is not voided. */
export function isJobLocked(job: ProcurementJob): boolean {
  return Object.values(job.issuedDocs).some((d) => Boolean(d && !d.isVoided));
}

/** A job is complete once its Delivery Note has been issued. */
export function isJobComplete(job: ProcurementJob): boolean {
  const dn = job.issuedDocs.DN;
  return Boolean(dn && !dn.isVoided);
}

/**
 * Removes earlier revisions from completed jobs. Returns the same array instance
 * when nothing changed so callers can skip a state update.
 */
export function purgeRevisionHistory(jobs: ProcurementJob[]): ProcurementJob[] {
  let changed = false;
  const next = jobs.map((job) => {
    if (!isJobComplete(job)) return job;
    let jobChanged = false;
    const issuedDocs = { ...job.issuedDocs };
    (Object.keys(issuedDocs) as (keyof typeof issuedDocs)[]).forEach((p) => {
      const doc = issuedDocs[p];
      if (doc && doc.history && doc.history.length > 0) {
        issuedDocs[p] = { ...doc, history: undefined };
        jobChanged = true;
      }
    });
    if (!jobChanged) return job;
    changed = true;
    return { ...job, issuedDocs };
  });
  return changed ? next : jobs;
}

const AUDIT_COLUMNS: { key: keyof import('../types/procurement').AuditEntry; label: string }[] = [
  { key: 'at', label: 'Date/Time (UTC)' },
  { key: 'docNumber', label: 'Document No.' },
  { key: 'prefix', label: 'Type' },
  { key: 'revision', label: 'Revision' },
  { key: 'jobName', label: 'Job' },
  { key: 'field', label: 'Field changed' },
  { key: 'from', label: 'Old value' },
  { key: 'to', label: 'New value' },
  { key: 'reason', label: 'Reason' },
];

/** RFC 4180 CSV. A UTF-8 BOM is added so Excel opens accented text correctly. */
export function auditLogToCsv(entries: import('../types/procurement').AuditEntry[]): string {
  const esc = (v: unknown) => {
    const s = v === null || v === undefined ? '' : String(v);
    return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [AUDIT_COLUMNS.map((c) => esc(c.label)).join(',')];
  for (const e of entries) {
    lines.push(AUDIT_COLUMNS.map((c) => esc(e[c.key])).join(','));
  }
  return '﻿' + lines.join('\r\n') + '\r\n';
}

export const AUDIT_EXPORT_COLUMNS = AUDIT_COLUMNS;
