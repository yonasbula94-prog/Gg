import { relations } from 'drizzle-orm';
import {
  boolean,
  doublePrecision,
  integer,
  jsonb,
  pgTable,
  text,
  timestamp,
  uniqueIndex,
} from 'drizzle-orm/pg-core';

export const users = pgTable('users', {
  uid: text('uid').primaryKey(),
  email: text('email').notNull().unique(),
  displayName: text('display_name').notNull().default(''),
  photoUrl: text('photo_url').notNull().default(''),
  isVendorAdmin: boolean('is_vendor_admin').notNull().default(false),
  createdAt: timestamp('created_at').defaultNow().notNull(),
  updatedAt: timestamp('updated_at').defaultNow().notNull(),
});

export const organizations = pgTable('organizations', {
  id: text('id').primaryKey(),
  name: text('name').notNull(),
  orgCode: text('org_code').notNull().unique(),
  companyProfileJson: jsonb('company_profile_json').notNull(),
  themeVariant: text('theme_variant').notNull().default('v1'),
  themeCustomJson: jsonb('theme_custom_json').notNull().default({}),
  dashboardPrefsJson: jsonb('dashboard_prefs_json').notNull().default({}),
  planTier: text('plan_tier').notNull().default('TRIAL'),
  seatLimit: integer('seat_limit').notNull().default(5),
  licenseExpiresAt: timestamp('license_expires_at').notNull(),
  activeLicenseId: text('active_license_id'),
  createdByUid: text('created_by_uid').references(() => users.uid, { onDelete: 'set null' }),
  createdAt: timestamp('created_at').defaultNow().notNull(),
  updatedAt: timestamp('updated_at').defaultNow().notNull(),
});

export const orgLicenses = pgTable('org_licenses', {
  id: text('id').primaryKey(),
  organizationId: text('organization_id').references(() => organizations.id, { onDelete: 'set null' }),
  targetOrgName: text('target_org_name').notNull(),
  planTier: text('plan_tier').notNull().default('BUSINESS'),
  seatLimit: integer('seat_limit').notNull().default(5),
  issuedAt: timestamp('issued_at').defaultNow().notNull(),
  expiresAt: timestamp('expires_at').notNull(),
  licenseKey: text('license_key').notNull().unique(),
  status: text('status').notNull().default('ISSUED'),
  issuedByUid: text('issued_by_uid').references(() => users.uid, { onDelete: 'set null' }),
  activatedByUid: text('activated_by_uid').references(() => users.uid, { onDelete: 'set null' }),
  activatedAt: timestamp('activated_at'),
  notes: text('notes').notNull().default(''),
});

export const orgMemberships = pgTable(
  'org_memberships',
  {
    id: text('id').primaryKey(),
    organizationId: text('organization_id')
      .references(() => organizations.id, { onDelete: 'cascade' })
      .notNull(),
    userUid: text('user_uid').references(() => users.uid, { onDelete: 'set null' }),
    invitedEmail: text('invited_email').notNull(),
    displayName: text('display_name').notNull().default(''),
    role: text('role').notNull().default('PROCUREMENT_OFFICER'),
    seatStatus: text('seat_status').notNull().default('ACTIVE'),
    invitedByUid: text('invited_by_uid').references(() => users.uid, { onDelete: 'set null' }),
    lastActiveAt: timestamp('last_active_at'),
    createdAt: timestamp('created_at').defaultNow().notNull(),
    updatedAt: timestamp('updated_at').defaultNow().notNull(),
  },
  (table) => [uniqueIndex('org_memberships_org_email_idx').on(table.organizationId, table.invitedEmail)]
);

export const orgSequences = pgTable('org_sequences', {
  id: text('id').primaryKey(),
  organizationId: text('organization_id')
    .references(() => organizations.id, { onDelete: 'cascade' })
    .notNull(),
  year: integer('year').notNull(),
  lastSequence: integer('last_sequence').notNull().default(0),
  updatedAt: timestamp('updated_at').defaultNow().notNull(),
});

export const procurementJobs = pgTable(
  'procurement_jobs',
  {
    id: text('id').primaryKey(),
    organizationId: text('organization_id')
      .references(() => organizations.id, { onDelete: 'cascade' })
      .notNull(),
    clientJobId: text('client_job_id').notNull(),
    jobName: text('job_name').notNull(),
    inquiryType: text('inquiry_type').notNull().default('RFQ'),
    jobSequence: integer('job_sequence'),
    jobSequenceYear: integer('job_sequence_year'),
    preparedBy: text('prepared_by').notNull().default(''),
    payloadJson: jsonb('payload_json').notNull(),
    version: integer('version').notNull().default(1),
    updatedByUid: text('updated_by_uid').references(() => users.uid, { onDelete: 'set null' }),
    updatedByName: text('updated_by_name').notNull().default(''),
    createdAt: timestamp('created_at').defaultNow().notNull(),
    updatedAt: timestamp('updated_at').defaultNow().notNull(),
  },
  (table) => [uniqueIndex('procurement_jobs_org_client_idx').on(table.organizationId, table.clientJobId)]
);

export const registryEntries = pgTable(
  'registry_entries',
  {
    id: text('id').primaryKey(),
    organizationId: text('organization_id')
      .references(() => organizations.id, { onDelete: 'cascade' })
      .notNull(),
    sequence: integer('sequence').notNull(),
    docNumber: text('doc_number').notNull(),
    prefix: text('prefix').notNull(),
    year: integer('year').notNull(),
    jobId: text('job_id').notNull(),
    jobName: text('job_name').notNull(),
    customer: text('customer').notNull(),
    counterpartyRole: text('counterparty_role').notNull().default('CLIENT'),
    rfqRef: text('rfq_ref').notNull().default(''),
    currency: text('currency').notNull().default('USD'),
    amount: doublePrecision('amount').notNull().default(0),
    issuedBy: text('issued_by').notNull(),
    issuedByUid: text('issued_by_uid').references(() => users.uid, { onDelete: 'set null' }),
    issuedAt: timestamp('issued_at').defaultNow().notNull(),
    status: text('status').notNull().default('ISSUED'),
    voidReason: text('void_reason'),
    revision: integer('revision').notNull().default(1),
  },
  (table) => [uniqueIndex('registry_entries_org_doc_idx').on(table.organizationId, table.docNumber)]
);

export const auditLogs = pgTable('audit_logs', {
  id: text('id').primaryKey(),
  organizationId: text('organization_id')
    .references(() => organizations.id, { onDelete: 'cascade' })
    .notNull(),
  actorUid: text('actor_uid').references(() => users.uid, { onDelete: 'set null' }),
  actorName: text('actor_name').notNull().default(''),
  eventType: text('event_type').notNull().default('DOCUMENT_AMENDMENT'),
  jobId: text('job_id').notNull().default(''),
  jobName: text('job_name').notNull().default(''),
  docNumber: text('doc_number').notNull().default(''),
  prefix: text('prefix').notNull().default(''),
  revision: integer('revision').notNull().default(1),
  reason: text('reason').notNull().default(''),
  field: text('field').notNull().default(''),
  fromVal: text('from_val').notNull().default(''),
  toVal: text('to_val').notNull().default(''),
  createdAt: timestamp('created_at').defaultNow().notNull(),
});

export const organizationsRelations = relations(organizations, ({ many }) => ({
  memberships: many(orgMemberships),
  licenses: many(orgLicenses),
  jobs: many(procurementJobs),
  registry: many(registryEntries),
  auditLogs: many(auditLogs),
}));

export const orgMembershipsRelations = relations(orgMemberships, ({ one }) => ({
  organization: one(organizations, {
    fields: [orgMemberships.organizationId],
    references: [organizations.id],
  }),
  user: one(users, {
    fields: [orgMemberships.userUid],
    references: [users.uid],
  }),
}));
