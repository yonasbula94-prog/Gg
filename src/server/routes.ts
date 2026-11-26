import crypto from 'node:crypto';
import express from 'express';
import { and, desc, eq, inArray, sql } from 'drizzle-orm';
import { db } from '../db/index.ts';
import {
  auditLogs,
  orgLicenses,
  orgMemberships,
  orgSequences,
  organizations,
  procurementJobs,
  registryEntries,
  users,
} from '../db/schema.ts';
import {
  requireAuth,
  requireOrgMember,
  requireVendorAdmin,
  type AuthRequest,
} from '../middleware/auth.ts';
import {
  createSignedOfflineLease,
  getPublicKeyPem,
  mintSignedLicenseKey,
  parseAndVerifyLicenseKey,
} from './licensing.ts';
import { buildExecutiveSnapshot, type AuditLogInput } from './executiveAggregator.ts';
import { streamHub } from './streamHub.ts';
import {
  calculateJob,
  getResolvedEndUser,
} from '../utils/calculations.ts';
import { COMPANY_INFO, DEFAULT_DASHBOARD_PREFS, DEFAULT_END_USER_ADDRESSES, END_USER_OPTIONS } from '../constants/themes.ts';
import type {
  CompanyProfile,
  DashboardPreferences,
  DocPrefix,
  IssuedDocRecord,
  IssuedRevision,
  ProcurementJob,
  RegistryEntry,
  ThemeCustomization,
  ThemeVariantId,
} from '../types/procurement.ts';
import type {
  ExecutiveSnapshot,
  LicensePlanTier,
  OrgRole,
  OrganizationSummary,
  PendingInvite,
  SeatStatus,
} from '../types/enterprise.ts';

const VALID_ROLES: OrgRole[] = [
  'ORG_ADMIN',
  'CEO_EXECUTIVE',
  'PROCUREMENT_OFFICER',
  'VIEWER',
];

const VALID_TIERS: LicensePlanTier[] = ['TRIAL', 'STARTER', 'BUSINESS', 'ENTERPRISE'];
const VALID_PREFIXES: DocPrefix[] = ['QT', 'PI', 'PO', 'PL', 'DN', 'CI'];

function createServerDefaultCompanyProfile(orgName?: string): CompanyProfile {
  return {
    ...COMPANY_INFO,
    name: orgName?.trim() || COMPANY_INFO.name,
    logoDataUrl: '',
    endUsers: [...END_USER_OPTIONS],
    endUserAddresses: { ...DEFAULT_END_USER_ADDRESSES },
  };
}

function createServerBlankJob(orgName: string, preparedBy: string): ProcurementJob {
  const now = new Date().toISOString();
  const defaultEndUser = END_USER_OPTIONS[0];
  return {
    id: `job-${Date.now()}-${crypto.randomBytes(2).toString('hex')}`,
    jobName: 'Procurement Job #1',
    inquiryType: 'RFQ',
    createdAt: now,
    updatedAt: now,
    preparedBy: preparedBy || 'Procurement Desk',
    version: 1,
    controls: {
      endUserSelect: defaultEndUser,
      customEndUser: '',
      endUserContact: '',
      endUserAddress: DEFAULT_END_USER_ADDRESSES[defaultEndUser] ?? '',
      currency: 'USD',
      incoterm: 'CIF',
      destination: 'Massawa Seaport',
      includeVat: true,
      vatRate: 0.05,
      paymentTerms: '50% Advance with PO, 50% Prior to Shipment',
      deliveryLeadTime: '4–6 Weeks from PO Confirmation',
      rfqRef: '',
      customerPoRef: '',
      supplierName: '',
      supplierAddress: '',
      supplierRef: '',
      shipDate: '',
      portOfLoading: 'Jebel Ali Port, UAE',
      deliveryDate: '',
      transportRef: '',
      receivedBy: '',
      remarks: 'All items supplied in original manufacturer export-grade packaging.',
      authorizedSignatory: `Authorized Signature — ${orgName}`,
      counterpartySignatory: 'Client / Supplier Acknowledgment',
    },
    supplierQuote: {
      supplierName: '',
      supplierContact: '',
      supplierAddress: '',
      supplierEmail: '',
      supplierQuoteRef: '',
      quoteDate: now.slice(0, 10),
      validUntil: '',
      currency: 'USD',
      incoterm: 'EXW',
      paymentTerms: '100% TT Prior to Dispatch',
      deliveryLeadTime: '2–3 Weeks from Order',
      notes: '',
    },
    workflow: {
      endUserPoConfirmed: false,
      supplierPaidConfirmed: false,
      supplierPaymentRef: '',
      supplierPackingListReceived: false,
      supplierPackingListRef: '',
    },
    freight: {
      ratePerCbm: 185,
      ratePerKg: 2.4,
      docChargePct: 0.1,
      selectedMarkupPct: 0.2,
      mode: 'air',
    },
    customColumns: [],
    items: Array.from({ length: 20 }, (_, i) => ({
      rowNumber: 17 + i,
      requestedDescription: '',
      offeredDescription: '',
      uom: '',
      qty: null,
      customFields: {},
      supplierUnitCost: null,
      boxes: null,
      boxLengthCm: null,
      boxWidthCm: null,
      boxHeightCm: null,
      kgPerBox: null,
      cartonNo: '',
      netWtKg: null,
      grossWtKg: null,
    })),
    issuedDocs: {},
  };
}

function makeOrgCode(name: string): string {
  const base = name
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 16) || 'ORG';
  return `${base}-${crypto.randomBytes(2).toString('hex').toUpperCase()}`;
}

function formatServerDocNumber(prefix: DocPrefix, sequence: number, year: number): string {
  return `${prefix}-GG-${String(sequence).padStart(3, '0')}-${year}`;
}

function mapRegistryRow(r: typeof registryEntries.$inferSelect): RegistryEntry {
  return {
    sequence: r.sequence,
    docNumber: r.docNumber,
    prefix: r.prefix as DocPrefix,
    year: r.year,
    jobId: r.jobId,
    jobName: r.jobName,
    customer: r.customer,
    counterpartyRole: r.counterpartyRole === 'SUPPLIER' ? 'SUPPLIER' : 'CLIENT',
    rfqRef: r.rfqRef,
    currency: r.currency,
    amount: r.amount,
    issuedBy: r.issuedBy,
    issuedAt: r.issuedAt.toISOString(),
    status: r.status === 'VOID' ? 'VOID' : 'ISSUED',
    voidReason: r.voidReason ?? undefined,
    revision: r.revision,
  };
}

async function getSeatCounts(orgId: string): Promise<{ active: number; invited: number }> {
  const rows = await db
    .select({ seatStatus: orgMemberships.seatStatus })
    .from(orgMemberships)
    .where(eq(orgMemberships.organizationId, orgId));
  let active = 0;
  let invited = 0;
  for (const r of rows) {
    if (r.seatStatus === 'ACTIVE') active += 1;
    else if (r.seatStatus === 'INVITED') invited += 1;
  }
  return { active, invited };
}

async function buildOrgSummaryForMember(
  org: typeof organizations.$inferSelect,
  member: typeof orgMemberships.$inferSelect
): Promise<OrganizationSummary> {
  const counts = await getSeatCounts(org.id);
  const nowMs = Date.now();
  const expMs = org.licenseExpiresAt.getTime();
  return {
    id: org.id,
    name: org.name,
    orgCode: org.orgCode,
    planTier: (VALID_TIERS.includes(org.planTier as LicensePlanTier)
      ? org.planTier
      : 'TRIAL') as LicensePlanTier,
    seatLimit: org.seatLimit,
    activeSeatsUsed: counts.active,
    invitedSeatsUsed: counts.invited,
    licenseExpiresAt: org.licenseExpiresAt.toISOString(),
    licenseExpired: expMs <= nowMs,
    daysUntilExpiry: Math.ceil((expMs - nowMs) / 86400000),
    role: member.role as OrgRole,
    seatStatus: member.seatStatus as SeatStatus,
  };
}

export async function computeOrgExecutiveSnapshot(orgId: string): Promise<ExecutiveSnapshot | null> {
  try {
    const [org] = await db.select().from(organizations).where(eq(organizations.id, orgId)).limit(1);
    if (!org) return null;

    const [jobRows, regRows, auditRows, seatCounts] = await Promise.all([
      db
        .select()
        .from(procurementJobs)
        .where(eq(procurementJobs.organizationId, orgId))
        .orderBy(desc(procurementJobs.updatedAt)),
      db
        .select()
        .from(registryEntries)
        .where(eq(registryEntries.organizationId, orgId))
        .orderBy(desc(registryEntries.issuedAt)),
      db
        .select()
        .from(auditLogs)
        .where(eq(auditLogs.organizationId, orgId))
        .orderBy(desc(auditLogs.createdAt))
        .limit(40),
      getSeatCounts(orgId),
    ]);

    const jobs: ProcurementJob[] = jobRows.map((row) => ({
      ...(row.payloadJson as ProcurementJob),
      id: row.clientJobId,
      version: row.version,
    }));

    const registry: RegistryEntry[] = regRows.map(mapRegistryRow);

    const recentAudit: AuditLogInput[] = auditRows.map((a) => ({
      id: a.id,
      at: a.createdAt.toISOString(),
      actorName: a.actorName,
      eventType: a.eventType,
      jobId: a.jobId,
      jobName: a.jobName,
      docNumber: a.docNumber,
      prefix: a.prefix,
      revision: a.revision,
      reason: a.reason,
      field: a.field,
    }));

    return buildExecutiveSnapshot({
      organizationId: org.id,
      organizationName: org.name,
      planTier: (VALID_TIERS.includes(org.planTier as LicensePlanTier)
        ? org.planTier
        : 'TRIAL') as LicensePlanTier,
      seatUsed: seatCounts.active,
      seatLimit: org.seatLimit,
      licenseExpiresAt: org.licenseExpiresAt,
      activeListeners: streamHub.getListenerCount(orgId),
      jobs,
      registry,
      recentAudit,
    });
  } catch (err) {
    console.error('Failed to compute executive snapshot:', err);
    throw new Error('Could not build executive dashboard snapshot.', { cause: err });
  }
}

async function broadcastExecutiveRefresh(orgId: string) {
  try {
    const snapshot = await computeOrgExecutiveSnapshot(orgId);
    if (snapshot) {
      streamHub.broadcastToOrg(orgId, 'executive:snapshot', snapshot);
    }
  } catch (err) {
    console.error('Failed to broadcast executive refresh:', err);
  }
}

async function recordAuditEvent(params: {
  organizationId: string;
  actorUid: string;
  actorName: string;
  eventType: string;
  jobId?: string;
  jobName?: string;
  docNumber?: string;
  prefix?: string;
  revision?: number;
  reason?: string;
  field?: string;
  fromVal?: string;
  toVal?: string;
}) {
  try {
    await db.insert(auditLogs).values({
      id: `aud-${Date.now().toString(36)}-${crypto.randomBytes(3).toString('hex')}`,
      organizationId: params.organizationId,
      actorUid: params.actorUid,
      actorName: params.actorName,
      eventType: params.eventType,
      jobId: params.jobId ?? '',
      jobName: params.jobName ?? '',
      docNumber: params.docNumber ?? '',
      prefix: params.prefix ?? '',
      revision: params.revision ?? 1,
      reason: params.reason ?? '',
      field: params.field ?? '',
      fromVal: params.fromVal ?? '',
      toVal: params.toVal ?? '',
    });
  } catch (err) {
    console.error('Audit log insert failed:', err);
  }
}

export function createApiRouter(): express.Router {
  const router = express.Router();

  // ==================== HEALTH ====================
  router.get('/health', (_req, res) => {
    res.json({ ok: true, timestamp: new Date().toISOString() });
  });

  // ==================== AUTH SESSION BOOTSTRAP ====================
  router.get('/auth/session', requireAuth, async (req: AuthRequest, res) => {
    try {
      const user = req.authUser!;
      // Fetch all memberships linked by userUid OR invited by email
      const memberRows = await db
        .select({
          member: orgMemberships,
          org: organizations,
        })
        .from(orgMemberships)
        .innerJoin(organizations, eq(orgMemberships.organizationId, organizations.id))
        .where(
          sql`${orgMemberships.userUid} = ${user.uid} OR ${orgMemberships.invitedEmail} = ${user.email}`
        )
        .orderBy(desc(orgMemberships.createdAt));

      const orgSummaries: OrganizationSummary[] = [];
      const pendingInvites: PendingInvite[] = [];

      for (const row of memberRows) {
        if (row.member.seatStatus === 'INVITED') {
          pendingInvites.push({
            membershipId: row.member.id,
            organizationId: row.org.id,
            organizationName: row.org.name,
            orgCode: row.org.orgCode,
            role: row.member.role as OrgRole,
            invitedAt: row.member.createdAt.toISOString(),
          });
        } else if (row.member.seatStatus === 'ACTIVE') {
          // Ensure userUid is linked if matched by email
          if (!row.member.userUid) {
            await db
              .update(orgMemberships)
              .set({ userUid: user.uid, displayName: user.displayName || row.member.displayName })
              .where(eq(orgMemberships.id, row.member.id));
          }
          orgSummaries.push(await buildOrgSummaryForMember(row.org, row.member));
        }
      }

      res.json({
        user,
        organizations: orgSummaries,
        pendingInvites,
      });
    } catch (err) {
      console.error('Session bootstrap failed:', err);
      res.status(500).json({ error: 'Could not load your account session.' });
    }
  });

  // ==================== CREATE ORGANIZATION (WITH OPTIONAL LICENSE KEY) ====================
  router.post('/orgs', requireAuth, async (req: AuthRequest, res) => {
    try {
      const user = req.authUser!;
      const rawName = typeof req.body?.name === 'string' ? req.body.name.trim() : '';
      const rawLicenseKey = typeof req.body?.licenseKey === 'string' ? req.body.licenseKey.trim() : '';

      if (!rawName || rawName.length < 2) {
        return res.status(400).json({ error: 'Organization name must be at least 2 characters.' });
      }

      let planTier: LicensePlanTier = 'TRIAL';
      let seatLimit = 5;
      let licenseExpiresAt = new Date(Date.now() + 14 * 86400 * 1000);
      let activeLicenseId: string | null = null;
      let dbLicenseRow: typeof orgLicenses.$inferSelect | null = null;

      if (rawLicenseKey) {
        const parsed = parseAndVerifyLicenseKey(rawLicenseKey);
        const [found] = await db
          .select()
          .from(orgLicenses)
          .where(eq(orgLicenses.licenseKey, rawLicenseKey))
          .limit(1);

        if (!found && !parsed.valid) {
          return res.status(400).json({
            error: parsed.error || 'License key is invalid or unrecognized.',
          });
        }
        if (found) {
          if (found.status === 'REVOKED') {
            return res.status(400).json({ error: 'This license key has been revoked.' });
          }
          if (found.status === 'ACTIVATED' && found.organizationId) {
            return res.status(400).json({ error: 'This license key is already activated on another organization.' });
          }
          if (found.expiresAt.getTime() <= Date.now()) {
            return res.status(400).json({ error: 'This license key has expired.' });
          }
          dbLicenseRow = found;
          planTier = found.planTier as LicensePlanTier;
          seatLimit = found.seatLimit;
          licenseExpiresAt = found.expiresAt;
          activeLicenseId = found.id;
        } else if (parsed.valid && parsed.payload) {
          planTier = parsed.payload.tier;
          seatLimit = parsed.payload.seats;
          licenseExpiresAt = new Date(parsed.payload.exp * 1000);
          activeLicenseId = parsed.payload.id;
        }
      }

      const orgId = `org_${crypto.randomBytes(6).toString('hex')}`;
      const orgCode = makeOrgCode(rawName);
      const membershipId = `mem_${crypto.randomBytes(6).toString('hex')}`;
      const companyProfile = createServerDefaultCompanyProfile(rawName);
      const initialJob = createServerBlankJob(rawName, user.displayName || user.email);

      const created = await db.transaction(async (tx) => {
        const [org] = await tx
          .insert(organizations)
          .values({
            id: orgId,
            name: rawName,
            orgCode,
            companyProfileJson: companyProfile,
            themeVariant: 'v1',
            themeCustomJson: {},
            dashboardPrefsJson: DEFAULT_DASHBOARD_PREFS,
            planTier,
            seatLimit,
            licenseExpiresAt,
            activeLicenseId,
            createdByUid: user.uid,
          })
          .returning();

        const [member] = await tx
          .insert(orgMemberships)
          .values({
            id: membershipId,
            organizationId: orgId,
            userUid: user.uid,
            invitedEmail: user.email,
            displayName: user.displayName || user.email.split('@')[0],
            role: 'ORG_ADMIN',
            seatStatus: 'ACTIVE',
            invitedByUid: user.uid,
            lastActiveAt: new Date(),
          })
          .returning();

        if (dbLicenseRow) {
          await tx
            .update(orgLicenses)
            .set({
              organizationId: orgId,
              status: 'ACTIVATED',
              activatedByUid: user.uid,
              activatedAt: new Date(),
            })
            .where(eq(orgLicenses.id, dbLicenseRow.id));
        }

        await tx.insert(procurementJobs).values({
          id: `${orgId}_${initialJob.id}`,
          organizationId: orgId,
          clientJobId: initialJob.id,
          jobName: initialJob.jobName,
          inquiryType: initialJob.inquiryType,
          preparedBy: initialJob.preparedBy,
          payloadJson: initialJob,
          version: 1,
          updatedByUid: user.uid,
          updatedByName: user.displayName || user.email,
        });

        return { org, member };
      });

      await recordAuditEvent({
        organizationId: orgId,
        actorUid: user.uid,
        actorName: user.displayName || user.email,
        eventType: 'LICENSE_ACTIVATED',
        reason: rawLicenseKey
          ? `Created organization with ${planTier} license (${seatLimit} seats)`
          : `Created organization on 14-day Trial (${seatLimit} seats)`,
      });

      const summary = await buildOrgSummaryForMember(created.org, created.member);
      res.status(201).json({ organization: summary });
    } catch (err) {
      console.error('Create organization failed:', err);
      res.status(500).json({ error: 'Could not create organization workspace.' });
    }
  });

  // ==================== ACCEPT PENDING INVITE ====================
  router.post('/orgs/invites/:membershipId/accept', requireAuth, async (req: AuthRequest, res) => {
    try {
      const user = req.authUser!;
      const membershipId = String(req.params.membershipId || '').trim();

      const result = await db.transaction(async (tx) => {
        const [member] = await tx
          .select()
          .from(orgMemberships)
          .where(eq(orgMemberships.id, membershipId))
          .limit(1);

        if (!member || member.invitedEmail !== user.email) {
          return { status: 404 as const, error: 'Invitation not found for your account.' };
        }
        if (member.seatStatus !== 'INVITED') {
          return { status: 400 as const, error: 'This invitation has already been processed.' };
        }

        const [org] = await tx
          .select()
          .from(organizations)
          .where(eq(organizations.id, member.organizationId))
          .limit(1);

        if (!org) {
          return { status: 404 as const, error: 'Organization no longer exists.' };
        }
        if (org.licenseExpiresAt.getTime() <= Date.now()) {
          return { status: 402 as const, error: 'Organization license has expired.' };
        }

        const activeRows = await tx
          .select({ id: orgMemberships.id })
          .from(orgMemberships)
          .where(
            and(
              eq(orgMemberships.organizationId, org.id),
              eq(orgMemberships.seatStatus, 'ACTIVE')
            )
          );

        if (activeRows.length >= org.seatLimit) {
          return {
            status: 409 as const,
            error: `All ${org.seatLimit} licensed seats for ${org.name} are currently occupied. Ask your Organization Admin to upgrade the license key.`,
          };
        }

        const [updatedMember] = await tx
          .update(orgMemberships)
          .set({
            userUid: user.uid,
            displayName: user.displayName || member.displayName || user.email.split('@')[0],
            seatStatus: 'ACTIVE',
            lastActiveAt: new Date(),
            updatedAt: new Date(),
          })
          .where(eq(orgMemberships.id, member.id))
          .returning();

        return { status: 200 as const, org, member: updatedMember };
      });

      if (result.status !== 200) {
        return res.status(result.status).json({ error: result.error });
      }

      await recordAuditEvent({
        organizationId: result.org.id,
        actorUid: user.uid,
        actorName: user.displayName || user.email,
        eventType: 'SEAT_CHANGED',
        reason: `${user.displayName || user.email} activated their ${result.member.role} seat`,
      });

      streamHub.broadcastToOrg(result.org.id, 'seat:updated', {
        memberId: result.member.id,
        seatStatus: 'ACTIVE',
      });
      await broadcastExecutiveRefresh(result.org.id);

      const summary = await buildOrgSummaryForMember(result.org, result.member);
      res.json({ organization: summary });
    } catch (err) {
      console.error('Accept invite failed:', err);
      res.status(500).json({ error: 'Could not accept invitation.' });
    }
  });

  // ==================== VENDOR LICENSE KEY CONSOLE ====================
  router.get('/vendor/licenses', requireAuth, requireVendorAdmin, async (_req, res) => {
    try {
      const rows = await db
        .select()
        .from(orgLicenses)
        .orderBy(desc(orgLicenses.issuedAt))
        .limit(100);

      res.json({
        publicKeyPem: getPublicKeyPem(),
        licenses: rows.map((l) => ({
          id: l.id,
          organizationId: l.organizationId,
          targetOrgName: l.targetOrgName,
          planTier: l.planTier as LicensePlanTier,
          seatLimit: l.seatLimit,
          issuedAt: l.issuedAt.toISOString(),
          expiresAt: l.expiresAt.toISOString(),
          licenseKey: l.licenseKey,
          status: l.status,
          activatedAt: l.activatedAt ? l.activatedAt.toISOString() : null,
          notes: l.notes,
        })),
      });
    } catch (err) {
      console.error('List vendor licenses failed:', err);
      res.status(500).json({ error: 'Could not load issued license ledger.' });
    }
  });

  router.post('/vendor/licenses', requireAuth, requireVendorAdmin, async (req: AuthRequest, res) => {
    try {
      const user = req.authUser!;
      const targetOrgName =
        typeof req.body?.targetOrgName === 'string' ? req.body.targetOrgName.trim() : '';
      const rawTier = typeof req.body?.planTier === 'string' ? req.body.planTier.toUpperCase() : 'BUSINESS';
      const planTier: LicensePlanTier = VALID_TIERS.includes(rawTier as LicensePlanTier)
        ? (rawTier as LicensePlanTier)
        : 'BUSINESS';
      const seatLimit = Number(req.body?.seatLimit);
      const validDays = Number(req.body?.validDays);
      const notes = typeof req.body?.notes === 'string' ? req.body.notes.trim().slice(0, 300) : '';

      if (!targetOrgName || targetOrgName.length < 2) {
        return res.status(400).json({ error: 'Target organization name is required.' });
      }
      if (!Number.isFinite(seatLimit) || seatLimit < 1 || seatLimit > 1000) {
        return res.status(400).json({ error: 'Seat limit must be between 1 and 1,000.' });
      }
      if (!Number.isFinite(validDays) || validDays < 7 || validDays > 3650) {
        return res.status(400).json({ error: 'Validity period must be between 7 and 3,650 days.' });
      }

      const minted = mintSignedLicenseKey({
        targetOrgName,
        planTier,
        seatLimit,
        validDays,
      });

      const [saved] = await db
        .insert(orgLicenses)
        .values({
          id: minted.licenseId,
          organizationId: null,
          targetOrgName,
          planTier,
          seatLimit: minted.payload.seats,
          issuedAt: minted.issuedAt,
          expiresAt: minted.expiresAt,
          licenseKey: minted.licenseKey,
          status: 'ISSUED',
          issuedByUid: user.uid,
          notes,
        })
        .returning();

      res.status(201).json({
        license: {
          id: saved.id,
          organizationId: saved.organizationId,
          targetOrgName: saved.targetOrgName,
          planTier: saved.planTier,
          seatLimit: saved.seatLimit,
          issuedAt: saved.issuedAt.toISOString(),
          expiresAt: saved.expiresAt.toISOString(),
          licenseKey: saved.licenseKey,
          status: saved.status,
          activatedAt: null,
          notes: saved.notes,
        },
      });
    } catch (err) {
      console.error('Mint license key failed:', err);
      res.status(500).json({ error: 'Could not generate signed license key.' });
    }
  });

  router.post(
    '/vendor/licenses/:licenseId/revoke',
    requireAuth,
    requireVendorAdmin,
    async (req: AuthRequest, res) => {
      try {
        const licenseId = String(req.params.licenseId || '').trim();
        const [updated] = await db
          .update(orgLicenses)
          .set({ status: 'REVOKED' })
          .where(eq(orgLicenses.id, licenseId))
          .returning();

        if (!updated) {
          return res.status(404).json({ error: 'License key not found.' });
        }
        res.json({ ok: true, id: updated.id, status: updated.status });
      } catch (err) {
        console.error('Revoke license failed:', err);
        res.status(500).json({ error: 'Could not revoke license key.' });
      }
    }
  );

  // ==================== WORKSPACE BOOTSTRAP & SETTINGS ====================
  router.get(
    '/orgs/:orgId/workspace',
    requireAuth,
    requireOrgMember({ allowWhenExpired: true }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;

        const [[org], [member], jobRows, regRows] = await Promise.all([
          db.select().from(organizations).where(eq(organizations.id, ctx.organizationId)).limit(1),
          db.select().from(orgMemberships).where(eq(orgMemberships.id, ctx.membershipId)).limit(1),
          db
            .select()
            .from(procurementJobs)
            .where(eq(procurementJobs.organizationId, ctx.organizationId))
            .orderBy(desc(procurementJobs.updatedAt)),
          db
            .select()
            .from(registryEntries)
            .where(eq(registryEntries.organizationId, ctx.organizationId))
            .orderBy(registryEntries.sequence),
        ]);

        if (!org || !member) {
          return res.status(404).json({ error: 'Organization workspace not found.' });
        }

        const jobs: ProcurementJob[] = jobRows.map((row) => ({
          ...(row.payloadJson as ProcurementJob),
          id: row.clientJobId,
          version: row.version,
        }));

        const registry: RegistryEntry[] = regRows.map(mapRegistryRow);

        const maxRegSeq = registry.reduce((max, e) => (e.sequence > max ? e.sequence : max), 0);
        const maxJobSeq = jobs.reduce(
          (max, j) => (j.jobSequence && j.jobSequence > max ? j.jobSequence : max),
          0
        );
        const seqRows = await db
          .select()
          .from(orgSequences)
          .where(eq(orgSequences.organizationId, ctx.organizationId));
        const maxStoredSeq = seqRows.reduce(
          (max, s) => (s.lastSequence > max ? s.lastSequence : max),
          0
        );
        const nextSequence = Math.max(maxRegSeq, maxJobSeq, maxStoredSeq) + 1;

        const orgSummary = await buildOrgSummaryForMember(org, member);
        const offlineLease = createSignedOfflineLease({
          orgId: org.id,
          orgName: org.name,
          uid: user.uid,
          email: user.email,
          role: ctx.role,
          seatLimit: org.seatLimit,
          planTier: orgSummary.planTier,
          licenseExpiresAt: org.licenseExpiresAt,
          graceDays: 7,
        });

        res.json({
          organization: orgSummary,
          companyProfile: org.companyProfileJson as CompanyProfile,
          themeVariant: (org.themeVariant === 'v3' ? 'v3' : 'v1') as ThemeVariantId,
          themeCustomization: (org.themeCustomJson || {}) as ThemeCustomization,
          dashboardPrefs: {
            ...DEFAULT_DASHBOARD_PREFS,
            ...((org.dashboardPrefsJson as Partial<DashboardPreferences>) || {}),
          },
          jobs,
          registry,
          nextSequence,
          offlineLease,
        });
      } catch (err) {
        console.error('Load workspace failed:', err);
        res.status(500).json({ error: 'Could not load organization workspace.' });
      }
    }
  );

  router.patch(
    '/orgs/:orgId/settings',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN', 'CEO_EXECUTIVE', 'PROCUREMENT_OFFICER'] }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const updates: Partial<typeof organizations.$inferInsert> = {
          updatedAt: new Date(),
        };

        if (req.body?.companyProfile && typeof req.body.companyProfile === 'object') {
          if (ctx.role !== 'ORG_ADMIN' && ctx.role !== 'CEO_EXECUTIVE') {
            return res
              .status(403)
              .json({ error: 'Only Organization Admins or CEOs can modify company branding.' });
          }
          updates.companyProfileJson = req.body.companyProfile;
        }
        if (req.body?.themeVariant === 'v1' || req.body?.themeVariant === 'v3') {
          updates.themeVariant = req.body.themeVariant;
        }
        if (req.body?.themeCustomization && typeof req.body.themeCustomization === 'object') {
          updates.themeCustomJson = req.body.themeCustomization;
        }
        if (req.body?.dashboardPrefs && typeof req.body.dashboardPrefs === 'object') {
          updates.dashboardPrefsJson = req.body.dashboardPrefs;
        }

        await db
          .update(organizations)
          .set(updates)
          .where(eq(organizations.id, ctx.organizationId));

        res.json({ ok: true });
      } catch (err) {
        console.error('Update org settings failed:', err);
        res.status(500).json({ error: 'Could not save organization settings.' });
      }
    }
  );

  // ==================== IMPORT EXISTING LOCAL WORKSPACE ====================
  router.post(
    '/orgs/:orgId/import-local',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN', 'CEO_EXECUTIVE', 'PROCUREMENT_OFFICER'] }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;
        const incomingJobs: ProcurementJob[] = Array.isArray(req.body?.jobs) ? req.body.jobs : [];
        const incomingRegistry: RegistryEntry[] = Array.isArray(req.body?.registry)
          ? req.body.registry
          : [];
        const incomingCompany =
          req.body?.companyProfile && typeof req.body.companyProfile === 'object'
            ? (req.body.companyProfile as CompanyProfile)
            : null;

        let importedJobsCount = 0;
        let importedRegistryCount = 0;

        await db.transaction(async (tx) => {
          for (const job of incomingJobs.slice(0, 200)) {
            if (!job || typeof job.id !== 'string' || !job.jobName) continue;
            const dbId = `${ctx.organizationId}_${job.id}`;
            await tx
              .insert(procurementJobs)
              .values({
                id: dbId,
                organizationId: ctx.organizationId,
                clientJobId: job.id,
                jobName: job.jobName,
                inquiryType: job.inquiryType === 'TENDER' ? 'TENDER' : 'RFQ',
                jobSequence: typeof job.jobSequence === 'number' ? job.jobSequence : null,
                jobSequenceYear:
                  typeof job.jobSequenceYear === 'number' ? job.jobSequenceYear : null,
                preparedBy: job.preparedBy || user.displayName || user.email,
                payloadJson: { ...job, version: 1 },
                version: 1,
                updatedByUid: user.uid,
                updatedByName: user.displayName || user.email,
              })
              .onConflictDoNothing();
            importedJobsCount += 1;
          }

          const maxSeqByYear = new Map<number, number>();
          for (const reg of incomingRegistry.slice(0, 1000)) {
            if (!reg || typeof reg.docNumber !== 'string' || typeof reg.sequence !== 'number') {
              continue;
            }
            const yr = Number.isFinite(reg.year) ? reg.year : new Date().getFullYear();
            maxSeqByYear.set(yr, Math.max(maxSeqByYear.get(yr) ?? 0, reg.sequence));
            await tx
              .insert(registryEntries)
              .values({
                id: `${ctx.organizationId}:${reg.docNumber}`,
                organizationId: ctx.organizationId,
                sequence: reg.sequence,
                docNumber: reg.docNumber,
                prefix: reg.prefix,
                year: yr,
                jobId: reg.jobId || `imported-${reg.sequence}`,
                jobName: reg.jobName || 'Imported Job',
                customer: reg.customer || 'Client',
                counterpartyRole: reg.counterpartyRole === 'SUPPLIER' ? 'SUPPLIER' : 'CLIENT',
                rfqRef: reg.rfqRef || '',
                currency: reg.currency || 'USD',
                amount: Number.isFinite(reg.amount) ? reg.amount : 0,
                issuedBy: reg.issuedBy || user.displayName || user.email,
                issuedByUid: user.uid,
                issuedAt: reg.issuedAt ? new Date(reg.issuedAt) : new Date(),
                status: reg.status === 'VOID' ? 'VOID' : 'ISSUED',
                voidReason: reg.voidReason ?? null,
                revision: reg.revision ?? 1,
              })
              .onConflictDoNothing();
            importedRegistryCount += 1;
          }

          for (const [yr, maxSeq] of maxSeqByYear.entries()) {
            const seqKey = `${ctx.organizationId}:${yr}`;
            await tx
              .insert(orgSequences)
              .values({
                id: seqKey,
                organizationId: ctx.organizationId,
                year: yr,
                lastSequence: maxSeq,
                updatedAt: new Date(),
              })
              .onConflictDoUpdate({
                target: orgSequences.id,
                set: {
                  lastSequence: sql`GREATEST(${orgSequences.lastSequence}, ${maxSeq})`,
                  updatedAt: new Date(),
                },
              });
          }

          if (incomingCompany && (ctx.role === 'ORG_ADMIN' || ctx.role === 'CEO_EXECUTIVE')) {
            await tx
              .update(organizations)
              .set({ companyProfileJson: incomingCompany, updatedAt: new Date() })
              .where(eq(organizations.id, ctx.organizationId));
          }
        });

        await recordAuditEvent({
          organizationId: ctx.organizationId,
          actorUid: user.uid,
          actorName: user.displayName || user.email,
          eventType: 'JOB_UPDATED',
          reason: `Migrated ${importedJobsCount} local jobs and ${importedRegistryCount} registry records into cloud organization`,
        });

        await broadcastExecutiveRefresh(ctx.organizationId);
        res.json({ ok: true, importedJobsCount, importedRegistryCount });
      } catch (err) {
        console.error('Local workspace import failed:', err);
        res.status(500).json({ error: 'Could not migrate local workspace data.' });
      }
    }
  );

  // ==================== ORGANIZATION SEATS & LICENSE ACTIVATION ====================
  router.get(
    '/orgs/:orgId/members',
    requireAuth,
    requireOrgMember({ allowWhenExpired: true }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const [members, licenses, [org]] = await Promise.all([
          db
            .select()
            .from(orgMemberships)
            .where(eq(orgMemberships.organizationId, ctx.organizationId))
            .orderBy(orgMemberships.createdAt),
          db
            .select()
            .from(orgLicenses)
            .where(eq(orgLicenses.organizationId, ctx.organizationId))
            .orderBy(desc(orgLicenses.activatedAt)),
          db.select().from(organizations).where(eq(organizations.id, ctx.organizationId)).limit(1),
        ]);

        const activeSeatsUsed = members.filter((m) => m.seatStatus === 'ACTIVE').length;
        const invitedSeatsUsed = members.filter((m) => m.seatStatus === 'INVITED').length;

        res.json({
          organization: org
            ? {
                id: org.id,
                name: org.name,
                orgCode: org.orgCode,
                planTier: org.planTier,
                seatLimit: org.seatLimit,
                activeSeatsUsed,
                invitedSeatsUsed,
                licenseExpiresAt: org.licenseExpiresAt.toISOString(),
                licenseExpired: org.licenseExpiresAt.getTime() <= Date.now(),
                daysUntilExpiry: Math.ceil((org.licenseExpiresAt.getTime() - Date.now()) / 86400000),
              }
            : null,
          members: members.map((m) => ({
            id: m.id,
            organizationId: m.organizationId,
            userUid: m.userUid,
            invitedEmail: m.invitedEmail,
            displayName: m.displayName,
            role: m.role as OrgRole,
            seatStatus: m.seatStatus as SeatStatus,
            lastActiveAt: m.lastActiveAt ? m.lastActiveAt.toISOString() : null,
            createdAt: m.createdAt.toISOString(),
          })),
          licenses: licenses.map((l) => ({
            id: l.id,
            organizationId: l.organizationId,
            targetOrgName: l.targetOrgName,
            planTier: l.planTier as LicensePlanTier,
            seatLimit: l.seatLimit,
            issuedAt: l.issuedAt.toISOString(),
            expiresAt: l.expiresAt.toISOString(),
            licenseKey: l.licenseKey,
            status: l.status,
            activatedAt: l.activatedAt ? l.activatedAt.toISOString() : null,
            notes: l.notes,
          })),
        });
      } catch (err) {
        console.error('List org members failed:', err);
        res.status(500).json({ error: 'Could not load organization members.' });
      }
    }
  );

  router.post(
    '/orgs/:orgId/license/activate',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN'], allowWhenExpired: true }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;
        const rawKey = typeof req.body?.licenseKey === 'string' ? req.body.licenseKey.trim() : '';
        if (!rawKey) {
          return res.status(400).json({ error: 'Please paste a signed GGPRO license key.' });
        }

        const parsed = parseAndVerifyLicenseKey(rawKey);
        const [found] = await db
          .select()
          .from(orgLicenses)
          .where(eq(orgLicenses.licenseKey, rawKey))
          .limit(1);

        if (!found && !parsed.valid) {
          return res.status(400).json({ error: parsed.error || 'Invalid license key.' });
        }

        let planTier: LicensePlanTier;
        let seatLimit: number;
        let expiresAt: Date;
        let licenseId: string;

        if (found) {
          if (found.status === 'REVOKED') {
            return res.status(400).json({ error: 'This license key has been revoked.' });
          }
          if (
            found.status === 'ACTIVATED' &&
            found.organizationId &&
            found.organizationId !== ctx.organizationId
          ) {
            return res
              .status(400)
              .json({ error: 'This license key is already activated on a different organization.' });
          }
          if (found.expiresAt.getTime() <= Date.now()) {
            return res.status(400).json({ error: 'This license key has expired.' });
          }
          planTier = found.planTier as LicensePlanTier;
          seatLimit = found.seatLimit;
          expiresAt = found.expiresAt;
          licenseId = found.id;
        } else {
          planTier = parsed.payload!.tier;
          seatLimit = parsed.payload!.seats;
          expiresAt = new Date(parsed.payload!.exp * 1000);
          licenseId = parsed.payload!.id;
        }

        await db.transaction(async (tx) => {
          await tx
            .update(organizations)
            .set({
              planTier,
              seatLimit,
              licenseExpiresAt: expiresAt,
              activeLicenseId: licenseId,
              updatedAt: new Date(),
            })
            .where(eq(organizations.id, ctx.organizationId));

          if (found) {
            await tx
              .update(orgLicenses)
              .set({
                organizationId: ctx.organizationId,
                status: 'ACTIVATED',
                activatedByUid: user.uid,
                activatedAt: new Date(),
              })
              .where(eq(orgLicenses.id, found.id));
          } else {
            await tx
              .insert(orgLicenses)
              .values({
                id: licenseId,
                organizationId: ctx.organizationId,
                targetOrgName: parsed.payload!.org,
                planTier,
                seatLimit,
                issuedAt: new Date(parsed.payload!.iat * 1000),
                expiresAt,
                licenseKey: rawKey,
                status: 'ACTIVATED',
                activatedByUid: user.uid,
                activatedAt: new Date(),
                notes: 'Activated via cryptographic Ed25519 signature verification',
              })
              .onConflictDoNothing();
          }
        });

        await recordAuditEvent({
          organizationId: ctx.organizationId,
          actorUid: user.uid,
          actorName: user.displayName || user.email,
          eventType: 'LICENSE_ACTIVATED',
          reason: `Applied ${planTier} license key (${seatLimit} seats, valid until ${expiresAt.toISOString().slice(0, 10)})`,
        });

        streamHub.broadcastToOrg(ctx.organizationId, 'seat:updated', {
          planTier,
          seatLimit,
          licenseExpiresAt: expiresAt.toISOString(),
        });
        await broadcastExecutiveRefresh(ctx.organizationId);

        res.json({
          ok: true,
          planTier,
          seatLimit,
          licenseExpiresAt: expiresAt.toISOString(),
          licenseId,
        });
      } catch (err) {
        console.error('Activate license failed:', err);
        res.status(500).json({ error: 'Could not activate license key.' });
      }
    }
  );

  router.post(
    '/orgs/:orgId/members/invite',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN'] }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;
        const email =
          typeof req.body?.email === 'string' ? req.body.email.trim().toLowerCase() : '';
        const displayName =
          typeof req.body?.displayName === 'string' ? req.body.displayName.trim().slice(0, 80) : '';
        const rawRole =
          typeof req.body?.role === 'string' ? req.body.role.trim().toUpperCase() : 'PROCUREMENT_OFFICER';
        const role: OrgRole = VALID_ROLES.includes(rawRole as OrgRole)
          ? (rawRole as OrgRole)
          : 'PROCUREMENT_OFFICER';

        if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
          return res.status(400).json({ error: 'Enter a valid work email address.' });
        }

        const outcome = await db.transaction(async (tx) => {
          const existing = await tx
            .select()
            .from(orgMemberships)
            .where(eq(orgMemberships.organizationId, ctx.organizationId));

          const duplicate = existing.find((m) => m.invitedEmail === email);
          if (duplicate) {
            return {
              status: 409 as const,
              error: `${email} is already listed in this organization (${duplicate.seatStatus}).`,
            };
          }

          const seatsReserved = existing.filter(
            (m) => m.seatStatus === 'ACTIVE' || m.seatStatus === 'INVITED'
          ).length;

          if (seatsReserved >= ctx.seatLimit) {
            return {
              status: 402 as const,
              error: `Seat limit reached (${seatsReserved} / ${ctx.seatLimit} seats used). Apply an upgraded license key or suspend an inactive seat first.`,
              code: 'SEAT_LIMIT_REACHED',
            };
          }

          const [knownUser] = await tx
            .select()
            .from(users)
            .where(eq(users.email, email))
            .limit(1);

          const [inserted] = await tx
            .insert(orgMemberships)
            .values({
              id: `mem_${crypto.randomBytes(6).toString('hex')}`,
              organizationId: ctx.organizationId,
              userUid: knownUser?.uid ?? null,
              invitedEmail: email,
              displayName: displayName || knownUser?.displayName || email.split('@')[0],
              role,
              seatStatus: knownUser ? 'ACTIVE' : 'INVITED',
              invitedByUid: user.uid,
              lastActiveAt: knownUser ? new Date() : null,
            })
            .returning();

          return { status: 201 as const, member: inserted };
        });

        if (outcome.status !== 201) {
          return res
            .status(outcome.status)
            .json({ error: outcome.error, code: (outcome as { code?: string }).code });
        }

        await recordAuditEvent({
          organizationId: ctx.organizationId,
          actorUid: user.uid,
          actorName: user.displayName || user.email,
          eventType: 'SEAT_CHANGED',
          reason: `Assigned ${role} seat to ${email} (${outcome.member.seatStatus})`,
        });

        streamHub.broadcastToOrg(ctx.organizationId, 'seat:updated', {
          memberId: outcome.member.id,
        });
        await broadcastExecutiveRefresh(ctx.organizationId);

        res.status(201).json({
          member: {
            id: outcome.member.id,
            organizationId: outcome.member.organizationId,
            userUid: outcome.member.userUid,
            invitedEmail: outcome.member.invitedEmail,
            displayName: outcome.member.displayName,
            role: outcome.member.role,
            seatStatus: outcome.member.seatStatus,
            lastActiveAt: outcome.member.lastActiveAt?.toISOString() ?? null,
            createdAt: outcome.member.createdAt.toISOString(),
          },
        });
      } catch (err) {
        console.error('Invite member failed:', err);
        res.status(500).json({ error: 'Could not invite team member.' });
      }
    }
  );

  router.patch(
    '/orgs/:orgId/members/:memberId',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN'] }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;
        const memberId = String(req.params.memberId || '').trim();
        const nextRole =
          typeof req.body?.role === 'string' && VALID_ROLES.includes(req.body.role as OrgRole)
            ? (req.body.role as OrgRole)
            : undefined;
        const nextStatus =
          req.body?.seatStatus === 'ACTIVE' || req.body?.seatStatus === 'SUSPENDED'
            ? (req.body.seatStatus as SeatStatus)
            : undefined;

        const outcome = await db.transaction(async (tx) => {
          const allMembers = await tx
            .select()
            .from(orgMemberships)
            .where(eq(orgMemberships.organizationId, ctx.organizationId));

          const target = allMembers.find((m) => m.id === memberId);
          if (!target) {
            return { status: 404 as const, error: 'Organization member not found.' };
          }
          if (target.id === ctx.membershipId && nextStatus === 'SUSPENDED') {
            return { status: 400 as const, error: 'You cannot suspend your own administrator seat.' };
          }
          if (target.id === ctx.membershipId && nextRole && nextRole !== 'ORG_ADMIN') {
            const otherAdmins = allMembers.filter(
              (m) => m.id !== target.id && m.role === 'ORG_ADMIN' && m.seatStatus === 'ACTIVE'
            );
            if (otherAdmins.length === 0) {
              return {
                status: 400 as const,
                error: 'Promote another Organization Admin before changing your own role.',
              };
            }
          }

          if (nextStatus === 'ACTIVE' && target.seatStatus === 'SUSPENDED') {
            const seatsUsed = allMembers.filter(
              (m) => m.seatStatus === 'ACTIVE' || m.seatStatus === 'INVITED'
            ).length;
            if (seatsUsed >= ctx.seatLimit) {
              return {
                status: 402 as const,
                error: `Cannot reactivate seat: all ${ctx.seatLimit} licensed seats are currently in use.`,
              };
            }
          }

          const resolvedStatus =
            nextStatus === 'ACTIVE' && !target.userUid ? 'INVITED' : nextStatus ?? target.seatStatus;

          const [updated] = await tx
            .update(orgMemberships)
            .set({
              role: nextRole ?? target.role,
              seatStatus: resolvedStatus,
              updatedAt: new Date(),
            })
            .where(eq(orgMemberships.id, target.id))
            .returning();

          return { status: 200 as const, member: updated, suspended: nextStatus === 'SUSPENDED' };
        });

        if (outcome.status !== 200) {
          return res.status(outcome.status).json({ error: outcome.error });
        }

        if (outcome.suspended && outcome.member.userUid) {
          streamHub.revokeUserStreams(
            ctx.organizationId,
            outcome.member.userUid,
            'Seat suspended by Organization Admin'
          );
        }

        await recordAuditEvent({
          organizationId: ctx.organizationId,
          actorUid: user.uid,
          actorName: user.displayName || user.email,
          eventType: 'SEAT_CHANGED',
          reason: `Updated ${outcome.member.invitedEmail} → ${outcome.member.role} (${outcome.member.seatStatus})`,
        });

        streamHub.broadcastToOrg(ctx.organizationId, 'seat:updated', {
          memberId: outcome.member.id,
        });
        await broadcastExecutiveRefresh(ctx.organizationId);

        res.json({
          member: {
            id: outcome.member.id,
            organizationId: outcome.member.organizationId,
            userUid: outcome.member.userUid,
            invitedEmail: outcome.member.invitedEmail,
            displayName: outcome.member.displayName,
            role: outcome.member.role,
            seatStatus: outcome.member.seatStatus,
            lastActiveAt: outcome.member.lastActiveAt?.toISOString() ?? null,
            createdAt: outcome.member.createdAt.toISOString(),
          },
        });
      } catch (err) {
        console.error('Update member failed:', err);
        res.status(500).json({ error: 'Could not update member seat.' });
      }
    }
  );

  router.delete(
    '/orgs/:orgId/members/:memberId',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN'] }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;
        const memberId = String(req.params.memberId || '').trim();

        if (memberId === ctx.membershipId) {
          return res.status(400).json({ error: 'You cannot remove your own active seat.' });
        }

        const [deleted] = await db
          .delete(orgMemberships)
          .where(
            and(
              eq(orgMemberships.id, memberId),
              eq(orgMemberships.organizationId, ctx.organizationId)
            )
          )
          .returning();

        if (!deleted) {
          return res.status(404).json({ error: 'Seat record not found.' });
        }

        if (deleted.userUid) {
          streamHub.revokeUserStreams(
            ctx.organizationId,
            deleted.userUid,
            'Removed from organization workspace'
          );
        }

        await recordAuditEvent({
          organizationId: ctx.organizationId,
          actorUid: user.uid,
          actorName: user.displayName || user.email,
          eventType: 'SEAT_CHANGED',
          reason: `Removed seat for ${deleted.invitedEmail}`,
        });

        streamHub.broadcastToOrg(ctx.organizationId, 'seat:updated', { removedId: memberId });
        await broadcastExecutiveRefresh(ctx.organizationId);
        res.json({ ok: true });
      } catch (err) {
        console.error('Remove member failed:', err);
        res.status(500).json({ error: 'Could not remove organization seat.' });
      }
    }
  );

  // ==================== PROCUREMENT JOBS CRUD + CONCURRENCY ====================
  router.put(
    '/orgs/:orgId/jobs/:clientJobId',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN', 'CEO_EXECUTIVE', 'PROCUREMENT_OFFICER'] }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;
        const clientJobId = String(req.params.clientJobId || '').trim();
        const incomingJob = req.body?.job as ProcurementJob | undefined;
        const baseVersion = typeof req.body?.baseVersion === 'number' ? req.body.baseVersion : undefined;
        const forceOverwrite = req.body?.forceOverwrite === true;
        const auditReason = typeof req.body?.auditReason === 'string' ? req.body.auditReason.trim() : '';

        if (!clientJobId || !incomingJob || typeof incomingJob !== 'object') {
          return res.status(400).json({ error: 'Valid job payload is required.' });
        }

        const dbId = `${ctx.organizationId}_${clientJobId}`;
        const [existing] = await db
          .select()
          .from(procurementJobs)
          .where(
            and(
              eq(procurementJobs.organizationId, ctx.organizationId),
              eq(procurementJobs.clientJobId, clientJobId)
            )
          )
          .limit(1);

        if (
          existing &&
          !forceOverwrite &&
          baseVersion !== undefined &&
          existing.version > baseVersion
        ) {
          return res.status(409).json({
            error: `This job was updated by ${existing.updatedByName || 'a colleague'} while you were editing.`,
            serverVersion: existing.version,
            serverJob: {
              ...(existing.payloadJson as ProcurementJob),
              id: existing.clientJobId,
              version: existing.version,
            },
          });
        }

        const nextVersion = existing ? existing.version + 1 : 1;
        const now = new Date();
        const normalizedPayload: ProcurementJob = {
          ...incomingJob,
          id: clientJobId,
          updatedAt: now.toISOString(),
          preparedBy: incomingJob.preparedBy || user.displayName || user.email,
          version: nextVersion,
        };

        await db
          .insert(procurementJobs)
          .values({
            id: dbId,
            organizationId: ctx.organizationId,
            clientJobId,
            jobName: normalizedPayload.jobName || 'Untitled Job',
            inquiryType: normalizedPayload.inquiryType === 'TENDER' ? 'TENDER' : 'RFQ',
            jobSequence: normalizedPayload.jobSequence ?? null,
            jobSequenceYear: normalizedPayload.jobSequenceYear ?? null,
            preparedBy: normalizedPayload.preparedBy,
            payloadJson: normalizedPayload,
            version: nextVersion,
            updatedByUid: user.uid,
            updatedByName: user.displayName || user.email,
            createdAt: now,
            updatedAt: now,
          })
          .onConflictDoUpdate({
            target: procurementJobs.id,
            set: {
              jobName: normalizedPayload.jobName || 'Untitled Job',
              inquiryType: normalizedPayload.inquiryType === 'TENDER' ? 'TENDER' : 'RFQ',
              jobSequence: normalizedPayload.jobSequence ?? null,
              jobSequenceYear: normalizedPayload.jobSequenceYear ?? null,
              preparedBy: normalizedPayload.preparedBy,
              payloadJson: normalizedPayload,
              version: nextVersion,
              updatedByUid: user.uid,
              updatedByName: user.displayName || user.email,
              updatedAt: now,
            },
          });

        if (!existing || auditReason) {
          await recordAuditEvent({
            organizationId: ctx.organizationId,
            actorUid: user.uid,
            actorName: user.displayName || user.email,
            eventType: existing ? 'JOB_UPDATED' : 'JOB_CREATED',
            jobId: clientJobId,
            jobName: normalizedPayload.jobName,
            reason: auditReason,
          });
        }

        streamHub.broadcastToOrg(ctx.organizationId, 'job:updated', {
          job: normalizedPayload,
          updatedByUid: user.uid,
          updatedByName: user.displayName || user.email,
        });
        await broadcastExecutiveRefresh(ctx.organizationId);

        res.json({ job: normalizedPayload, version: nextVersion });
      } catch (err) {
        console.error('Save job failed:', err);
        res.status(500).json({ error: 'Could not save job to server.' });
      }
    }
  );

  router.delete(
    '/orgs/:orgId/jobs/:clientJobId',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN', 'CEO_EXECUTIVE', 'PROCUREMENT_OFFICER'] }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;
        const clientJobId = String(req.params.clientJobId || '').trim();

        const [deleted] = await db
          .delete(procurementJobs)
          .where(
            and(
              eq(procurementJobs.organizationId, ctx.organizationId),
              eq(procurementJobs.clientJobId, clientJobId)
            )
          )
          .returning();

        if (!deleted) {
          return res.status(404).json({ error: 'Job not found.' });
        }

        await recordAuditEvent({
          organizationId: ctx.organizationId,
          actorUid: user.uid,
          actorName: user.displayName || user.email,
          eventType: 'JOB_DELETED',
          jobId: clientJobId,
          jobName: deleted.jobName,
        });

        streamHub.broadcastToOrg(ctx.organizationId, 'job:deleted', {
          jobId: clientJobId,
          deletedByUid: user.uid,
        });
        await broadcastExecutiveRefresh(ctx.organizationId);

        res.json({ ok: true });
      } catch (err) {
        console.error('Delete job failed:', err);
        res.status(500).json({ error: 'Could not delete job.' });
      }
    }
  );

  // ==================== ATOMIC DOCUMENT ISSUANCE (MASTER REGISTRY) ====================
  router.post(
    '/orgs/:orgId/jobs/:clientJobId/issue',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN', 'CEO_EXECUTIVE', 'PROCUREMENT_OFFICER'] }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;
        const clientJobId = String(req.params.clientJobId || '').trim();
        const prefix = String(req.body?.prefix || '').toUpperCase() as DocPrefix;
        const incomingJob = req.body?.job as ProcurementJob | undefined;

        if (!VALID_PREFIXES.includes(prefix)) {
          return res.status(400).json({ error: 'Invalid trade document prefix.' });
        }

        const result = await db.transaction(async (tx) => {
          const [jobRow] = await tx
            .select()
            .from(procurementJobs)
            .where(
              and(
                eq(procurementJobs.organizationId, ctx.organizationId),
                eq(procurementJobs.clientJobId, clientJobId)
              )
            )
            .limit(1);

          const baseJob: ProcurementJob | null = incomingJob
            ? { ...incomingJob, id: clientJobId }
            : jobRow
            ? { ...(jobRow.payloadJson as ProcurementJob), id: clientJobId }
            : null;

          if (!baseJob) {
            return { status: 404 as const, error: 'Job not found.' };
          }

          const existingIssued = baseJob.issuedDocs?.[prefix];
          if (existingIssued && !existingIssued.isVoided) {
            return {
              status: 409 as const,
              error: `${prefix} is already issued as ${existingIssued.docNumber}.`,
            };
          }

          const now = new Date();
          const year =
            typeof baseJob.jobSequenceYear === 'number' && baseJob.jobSequenceYear > 2000
              ? baseJob.jobSequenceYear
              : now.getFullYear();

          let seq =
            typeof baseJob.jobSequence === 'number' && baseJob.jobSequence > 0
              ? baseJob.jobSequence
              : 0;

          if (seq === 0) {
            // Atomically claim the next sequence number across all concurrent workstations
            const seqRowId = `${ctx.organizationId}:${year}`;
            await tx
              .insert(orgSequences)
              .values({
                id: seqRowId,
                organizationId: ctx.organizationId,
                year,
                lastSequence: 0,
                updatedAt: now,
              })
              .onConflictDoNothing();

            const [lockedSeq] = await tx
              .select()
              .from(orgSequences)
              .where(eq(orgSequences.id, seqRowId))
              .for('update');

            const regMaxResult = await tx
              .select({
                maxSeq: sql<number>`COALESCE(MAX(${registryEntries.sequence}), 0)`,
              })
              .from(registryEntries)
              .where(eq(registryEntries.organizationId, ctx.organizationId));

            const currentMax = Math.max(
              lockedSeq?.lastSequence ?? 0,
              Number(regMaxResult[0]?.maxSeq ?? 0)
            );
            seq = currentMax + 1;

            await tx
              .update(orgSequences)
              .set({ lastSequence: seq, updatedAt: now })
              .where(eq(orgSequences.id, seqRowId));
          }

          const docNumber = formatServerDocNumber(prefix, seq, year);
          const calc = calculateJob(baseJob);
          const actorLabel = user.displayName || baseJob.preparedBy || user.email;
          const counterpartyRole = prefix === 'PO' ? 'SUPPLIER' : 'CLIENT';
          const partyName =
            prefix === 'PO'
              ? baseJob.controls.supplierName || 'Supplier'
              : getResolvedEndUser(baseJob) || 'Client';
          const docAmount =
            prefix === 'PO'
              ? calc.totalSupplierCost
              : prefix === 'PL' || prefix === 'DN'
              ? 0
              : calc.grandTotal;

          let nextControls = baseJob.controls;
          let nextItems = baseJob.items;
          if (prefix === 'PL') {
            nextControls = {
              ...baseJob.controls,
              deliveryDate:
                baseJob.controls.deliveryDate.trim() ||
                baseJob.controls.shipDate ||
                now.toISOString().slice(0, 10),
            };
            nextItems = baseJob.items.map((item, idx) => {
              const comp = calc.rows[idx];
              if (!comp || comp.isEmpty) return item;
              return { ...item, grossWtKg: item.grossWtKg ?? comp.grossKgCalc };
            });
          }

          const issuedRecord: IssuedDocRecord = {
            prefix,
            docNumber,
            sequence: seq,
            issuedAt: now.toISOString(),
            issuedBy: actorLabel,
            frozenSubtotal: calc.goodsSubtotal,
            frozenVat: calc.vatAmount,
            frozenFreight: calc.customerFreight,
            frozenGrandTotal: calc.grandTotal,
            revision: 1,
          };

          const nextVersion = (jobRow?.version ?? baseJob.version ?? 0) + 1;
          const updatedJob: ProcurementJob = {
            ...baseJob,
            updatedAt: now.toISOString(),
            jobSequence: seq,
            jobSequenceYear: year,
            controls: nextControls,
            items: nextItems,
            version: nextVersion,
            issuedDocs: {
              ...baseJob.issuedDocs,
              [prefix]: issuedRecord,
            },
          };

          const dbJobId = `${ctx.organizationId}_${clientJobId}`;
          await tx
            .insert(procurementJobs)
            .values({
              id: dbJobId,
              organizationId: ctx.organizationId,
              clientJobId,
              jobName: updatedJob.jobName,
              inquiryType: updatedJob.inquiryType,
              jobSequence: seq,
              jobSequenceYear: year,
              preparedBy: updatedJob.preparedBy || actorLabel,
              payloadJson: updatedJob,
              version: nextVersion,
              updatedByUid: user.uid,
              updatedByName: actorLabel,
              createdAt: now,
              updatedAt: now,
            })
            .onConflictDoUpdate({
              target: procurementJobs.id,
              set: {
                jobName: updatedJob.jobName,
                inquiryType: updatedJob.inquiryType,
                jobSequence: seq,
                jobSequenceYear: year,
                payloadJson: updatedJob,
                version: nextVersion,
                updatedByUid: user.uid,
                updatedByName: actorLabel,
                updatedAt: now,
              },
            });

          const [regRow] = await tx
            .insert(registryEntries)
            .values({
              id: `${ctx.organizationId}:${docNumber}`,
              organizationId: ctx.organizationId,
              sequence: seq,
              docNumber,
              prefix,
              year,
              jobId: clientJobId,
              jobName: updatedJob.jobName,
              customer: partyName,
              counterpartyRole,
              rfqRef: updatedJob.controls.rfqRef || '',
              currency: updatedJob.controls.currency || 'USD',
              amount: docAmount,
              issuedBy: actorLabel,
              issuedByUid: user.uid,
              issuedAt: now,
              status: 'ISSUED',
              revision: 1,
            })
            .returning();

          return {
            status: 200 as const,
            job: updatedJob,
            registryEntry: mapRegistryRow(regRow),
            sequence: seq,
            docNumber,
          };
        });

        if (result.status !== 200) {
          return res.status(result.status).json({ error: result.error });
        }

        await recordAuditEvent({
          organizationId: ctx.organizationId,
          actorUid: user.uid,
          actorName: user.displayName || user.email,
          eventType: 'DOCUMENT_ISSUED',
          jobId: clientJobId,
          jobName: result.job.jobName,
          docNumber: result.docNumber,
          prefix,
          revision: 1,
        });

        streamHub.broadcastToOrg(ctx.organizationId, 'doc:issued', {
          job: result.job,
          registryEntry: result.registryEntry,
        });
        await broadcastExecutiveRefresh(ctx.organizationId);

        res.json({
          job: result.job,
          registryEntry: result.registryEntry,
          sequence: result.sequence,
          docNumber: result.docNumber,
        });
      } catch (err) {
        console.error('Issue document failed:', err);
        res.status(500).json({ error: 'Could not issue document and lock sequence number.' });
      }
    }
  );

  // ==================== SAVE ISSUED DOCUMENT AMENDMENT ====================
  router.post(
    '/orgs/:orgId/jobs/:clientJobId/amend',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN', 'CEO_EXECUTIVE', 'PROCUREMENT_OFFICER'] }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;
        const clientJobId = String(req.params.clientJobId || '').trim();
        const prefix = String(req.body?.prefix || '').toUpperCase() as DocPrefix;
        const reason = typeof req.body?.reason === 'string' ? req.body.reason.trim() : '';
        const incomingJob = req.body?.job as ProcurementJob | undefined;
        const changes = Array.isArray(req.body?.changes)
          ? (req.body.changes as { field: string; from: string; to: string }[])
          : [];

        if (!VALID_PREFIXES.includes(prefix) || !incomingJob) {
          return res.status(400).json({ error: 'Invalid amendment request.' });
        }

        const doc = incomingJob.issuedDocs?.[prefix];
        if (!doc || doc.isVoided) {
          return res.status(400).json({ error: 'Cannot amend an unissued or voided document.' });
        }

        const now = new Date();
        const previousRev = doc.revision ?? 1;
        const newRev = previousRev + 1;
        const newCalc = calculateJob(incomingJob);
        const docAmount =
          prefix === 'PO'
            ? newCalc.totalSupplierCost
            : prefix === 'PL' || prefix === 'DN'
            ? 0
            : newCalc.grandTotal;

        const superseded: IssuedRevision = {
          revision: previousRev,
          docNumber: doc.docNumber,
          supersededAt: now.toISOString(),
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

        const [existingRow] = await db
          .select()
          .from(procurementJobs)
          .where(
            and(
              eq(procurementJobs.organizationId, ctx.organizationId),
              eq(procurementJobs.clientJobId, clientJobId)
            )
          )
          .limit(1);

        const nextVersion = (existingRow?.version ?? incomingJob.version ?? 0) + 1;
        const updatedJob: ProcurementJob = {
          ...incomingJob,
          id: clientJobId,
          updatedAt: now.toISOString(),
          version: nextVersion,
          issuedDocs: {
            ...incomingJob.issuedDocs,
            [prefix]: nextDoc,
          },
        };

        await db.transaction(async (tx) => {
          await tx
            .update(procurementJobs)
            .set({
              jobName: updatedJob.jobName,
              payloadJson: updatedJob,
              version: nextVersion,
              updatedByUid: user.uid,
              updatedByName: user.displayName || user.email,
              updatedAt: now,
            })
            .where(
              and(
                eq(procurementJobs.organizationId, ctx.organizationId),
                eq(procurementJobs.clientJobId, clientJobId)
              )
            );

          await tx
            .update(registryEntries)
            .set({ amount: docAmount, revision: newRev })
            .where(
              and(
                eq(registryEntries.organizationId, ctx.organizationId),
                eq(registryEntries.docNumber, doc.docNumber)
              )
            );

          for (const ch of changes) {
            await tx.insert(auditLogs).values({
              id: `aud-${Date.now().toString(36)}-${crypto.randomBytes(3).toString('hex')}`,
              organizationId: ctx.organizationId,
              actorUid: user.uid,
              actorName: user.displayName || user.email,
              eventType: 'DOCUMENT_AMENDMENT',
              jobId: clientJobId,
              jobName: updatedJob.jobName,
              docNumber: doc.docNumber,
              prefix,
              revision: newRev,
              reason,
              field: String(ch.field || ''),
              fromVal: String(ch.from || ''),
              toVal: String(ch.to || ''),
            });
          }
        });

        streamHub.broadcastToOrg(ctx.organizationId, 'job:updated', {
          job: updatedJob,
          updatedByUid: user.uid,
          updatedByName: user.displayName || user.email,
        });
        await broadcastExecutiveRefresh(ctx.organizationId);

        res.json({ job: updatedJob, revision: newRev, docNumber: doc.docNumber, amount: docAmount });
      } catch (err) {
        console.error('Save revision failed:', err);
        res.status(500).json({ error: 'Could not save document revision.' });
      }
    }
  );

  // ==================== VOID REGISTRY ENTRY ====================
  router.post(
    '/orgs/:orgId/registry/:docNumber/void',
    requireAuth,
    requireOrgMember({ allowedRoles: ['ORG_ADMIN', 'CEO_EXECUTIVE'] }),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const user = req.authUser!;
        const docNumber = decodeURIComponent(String(req.params.docNumber || '').trim());
        const reason = typeof req.body?.reason === 'string' ? req.body.reason.trim() : '';

        if (!reason) {
          return res.status(400).json({ error: 'A void reason is required for audit compliance.' });
        }

        const [updatedReg] = await db
          .update(registryEntries)
          .set({ status: 'VOID', voidReason: reason })
          .where(
            and(
              eq(registryEntries.organizationId, ctx.organizationId),
              eq(registryEntries.docNumber, docNumber)
            )
          )
          .returning();

        if (!updatedReg) {
          return res.status(404).json({ error: 'Registry entry not found.' });
        }

        // Also mark voided inside the linked job payload
        const [jobRow] = await db
          .select()
          .from(procurementJobs)
          .where(
            and(
              eq(procurementJobs.organizationId, ctx.organizationId),
              eq(procurementJobs.clientJobId, updatedReg.jobId)
            )
          )
          .limit(1);

        if (jobRow) {
          const payload = jobRow.payloadJson as ProcurementJob;
          const prefix = updatedReg.prefix as DocPrefix;
          if (payload.issuedDocs?.[prefix]?.docNumber === docNumber) {
            payload.issuedDocs[prefix] = {
              ...payload.issuedDocs[prefix]!,
              isVoided: true,
            };
            await db
              .update(procurementJobs)
              .set({
                payloadJson: { ...payload, version: jobRow.version + 1 },
                version: jobRow.version + 1,
                updatedAt: new Date(),
              })
              .where(eq(procurementJobs.id, jobRow.id));
          }
        }

        await recordAuditEvent({
          organizationId: ctx.organizationId,
          actorUid: user.uid,
          actorName: user.displayName || user.email,
          eventType: 'DOCUMENT_VOIDED',
          jobId: updatedReg.jobId,
          jobName: updatedReg.jobName,
          docNumber,
          prefix: updatedReg.prefix,
          reason,
        });

        streamHub.broadcastToOrg(ctx.organizationId, 'registry:updated', {
          docNumber,
          status: 'VOID',
          voidReason: reason,
        });
        await broadcastExecutiveRefresh(ctx.organizationId);

        res.json({ ok: true, entry: mapRegistryRow(updatedReg) });
      } catch (err) {
        console.error('Void registry entry failed:', err);
        res.status(500).json({ error: 'Could not void registry entry.' });
      }
    }
  );

  // ==================== AUDIT LOG EXPORT ====================
  router.get(
    '/orgs/:orgId/audit',
    requireAuth,
    requireOrgMember(),
    async (req: AuthRequest, res) => {
      try {
        const ctx = req.orgContext!;
        const rows = await db
          .select()
          .from(auditLogs)
          .where(eq(auditLogs.organizationId, ctx.organizationId))
          .orderBy(desc(auditLogs.createdAt))
          .limit(500);

        res.json({
          entries: rows.map((a) => ({
            id: a.id,
            at: a.createdAt.toISOString(),
            actorName: a.actorName,
            eventType: a.eventType,
            jobId: a.jobId,
            jobName: a.jobName,
            docNumber: a.docNumber,
            prefix: (a.prefix || 'QT') as DocPrefix,
            revision: a.revision,
            reason: a.reason,
            field: a.field,
            from: a.fromVal,
            to: a.toVal,
          })),
        });
      } catch (err) {
        console.error('Fetch audit log failed:', err);
        res.status(500).json({ error: 'Could not load audit log.' });
      }
    }
  );

  // ==================== CEO LIVE EXECUTIVE DASHBOARD & STREAMING ====================
  router.get(
    '/orgs/:orgId/executive/snapshot',
    requireAuth,
    requireOrgMember(),
    async (req: AuthRequest, res) => {
      try {
        const snapshot = await computeOrgExecutiveSnapshot(req.orgContext!.organizationId);
        if (!snapshot) {
          return res.status(404).json({ error: 'Organization not found.' });
        }
        res.json(snapshot);
      } catch (err) {
        console.error('Executive snapshot route failed:', err);
        res.status(500).json({ error: 'Could not compute executive dashboard snapshot.' });
      }
    }
  );

  router.post(
    '/orgs/:orgId/stream-ticket',
    requireAuth,
    requireOrgMember(),
    async (req: AuthRequest, res) => {
      const ctx = req.orgContext!;
      const user = req.authUser!;
      const issued = streamHub.createStreamTicket({
        orgId: ctx.organizationId,
        uid: user.uid,
        role: ctx.role,
        displayName: user.displayName || user.email,
      });

      res.json({
        ticket: issued.ticket,
        expiresInSeconds: issued.expiresInSeconds,
        sseUrl: `/api/orgs/${encodeURIComponent(ctx.organizationId)}/executive/stream?ticket=${encodeURIComponent(issued.ticket)}`,
        wsPath: '/ws',
      });
    }
  );

  router.get('/orgs/:orgId/executive/stream', async (req, res) => {
    const orgId = String(req.params.orgId || '').trim();
    const ticketParam = typeof req.query.ticket === 'string' ? req.query.ticket.trim() : '';

    if (!orgId || !ticketParam) {
      return res.status(401).json({ error: 'Valid stream ticket is required.' });
    }

    const ticket = streamHub.consumeStreamTicket(ticketParam, orgId);
    if (!ticket) {
      return res.status(401).json({ error: 'Stream ticket is invalid or expired.' });
    }

    // Re-verify seat & license right before opening stream
    try {
      const rows = await db
        .select({ member: orgMemberships, org: organizations })
        .from(orgMemberships)
        .innerJoin(organizations, eq(orgMemberships.organizationId, organizations.id))
        .where(
          and(
            eq(orgMemberships.organizationId, orgId),
            eq(orgMemberships.userUid, ticket.uid),
            eq(orgMemberships.seatStatus, 'ACTIVE')
          )
        )
        .limit(1);

      const hit = rows[0];
      if (!hit || hit.org.licenseExpiresAt.getTime() <= Date.now()) {
        return res.status(403).json({ error: 'Active licensed organization seat is required.' });
      }

      res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
      res.setHeader('Cache-Control', 'no-cache, no-transform');
      res.setHeader('Connection', 'keep-alive');
      res.setHeader('X-Accel-Buffering', 'no');
      res.flushHeaders?.();

      const client = streamHub.registerSseClient({
        orgId,
        uid: ticket.uid,
        role: ticket.role,
        res,
      });

      const snapshot = await computeOrgExecutiveSnapshot(orgId);
      if (snapshot) {
        res.write(
          `event: executive:snapshot\ndata: ${JSON.stringify({
            type: 'executive:snapshot',
            orgId,
            at: new Date().toISOString(),
            payload: snapshot,
          })}\n\n`
        );
      }

      req.on('close', () => {
        streamHub.unregisterSseClient(client);
      });
    } catch (err) {
      console.error('SSE stream setup failed:', err);
      if (!res.headersSent) {
        res.status(500).json({ error: 'Could not start executive live stream.' });
      } else {
        res.end();
      }
    }
  });

  // Unused import guard for inArray if needed later
  void inArray;

  return router;
}
