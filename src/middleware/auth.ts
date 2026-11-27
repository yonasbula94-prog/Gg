import type { NextFunction, Request, Response } from 'express';
import { and, count, eq } from 'drizzle-orm';
import { adminAuth } from '../lib/firebase-admin.ts';
import { db } from '../db/index.ts';
import { orgMemberships, organizations, users } from '../db/schema.ts';
import type { OrgRole, SeatStatus } from '../types/enterprise.ts';

export interface AuthenticatedUser {
  uid: string;
  email: string;
  displayName: string;
  photoUrl: string;
  isVendorAdmin: boolean;
}

export interface OrgContext {
  organizationId: string;
  organizationName: string;
  orgCode: string;
  planTier: string;
  seatLimit: number;
  licenseExpiresAt: Date;
  licenseExpired: boolean;
  membershipId: string;
  role: OrgRole;
  seatStatus: SeatStatus;
}

export interface AuthRequest extends Request {
  authUser?: AuthenticatedUser;
  orgContext?: OrgContext;
}

function getConfiguredVendorEmails(): Set<string> {
  const configured = (process.env.VENDOR_ADMIN_EMAILS || 'yonasbula94@gmail.com')
    .split(',')
    .map((e) => e.trim().toLowerCase())
    .filter(Boolean);
  return new Set(configured);
}

async function upsertUserRecord(params: {
  uid: string;
  email: string;
  displayName: string;
  photoUrl: string;
}): Promise<AuthenticatedUser> {
  const emailLower = params.email.trim().toLowerCase();
  const vendorEmails = getConfiguredVendorEmails();

  let shouldBeVendorAdmin = vendorEmails.has(emailLower);
  if (!shouldBeVendorAdmin) {
    const existingVendor = await db
      .select({ total: count() })
      .from(users)
      .where(eq(users.isVendorAdmin, true));
    if ((existingVendor[0]?.total ?? 0) === 0) {
      shouldBeVendorAdmin = true;
    }
  }

  const now = new Date();
  const [record] = await db
    .insert(users)
    .values({
      uid: params.uid,
      email: emailLower,
      displayName: params.displayName || emailLower.split('@')[0],
      photoUrl: params.photoUrl || '',
      isVendorAdmin: shouldBeVendorAdmin,
      createdAt: now,
      updatedAt: now,
    })
    .onConflictDoUpdate({
      target: users.uid,
      set: {
        email: emailLower,
        displayName: params.displayName || emailLower.split('@')[0],
        photoUrl: params.photoUrl || '',
        updatedAt: now,
      },
    })
    .returning();

  // Ensure configured vendor email is always promoted
  if (vendorEmails.has(emailLower) && !record.isVendorAdmin) {
    const [updated] = await db
      .update(users)
      .set({ isVendorAdmin: true, updatedAt: now })
      .where(eq(users.uid, params.uid))
      .returning();
    return {
      uid: updated.uid,
      email: updated.email,
      displayName: updated.displayName,
      photoUrl: updated.photoUrl,
      isVendorAdmin: updated.isVendorAdmin,
    };
  }

  return {
    uid: record.uid,
    email: record.email,
    displayName: record.displayName,
    photoUrl: record.photoUrl,
    isVendorAdmin: record.isVendorAdmin,
  };
}

export const requireAuth = async (req: AuthRequest, res: Response, next: NextFunction) => {
  const customHeader = req.headers['x-firebase-auth'];
  const customToken = typeof customHeader === 'string' ? customHeader.trim() : '';

  const authHeader = typeof req.headers.authorization === 'string' ? req.headers.authorization : '';
  const bearerToken = authHeader.startsWith('Bearer ')
    ? authHeader.slice('Bearer '.length).trim()
    : '';

  const token = customToken || bearerToken;
  if (!token) {
    return res.status(401).json({ error: 'Authentication required. Please sign in.' });
  }

  let decoded;
  try {
    decoded = await adminAuth.verifyIdToken(token);
  } catch (err) {
    console.error('Auth token verification error:', err);
    return res.status(401).json({ error: 'Session expired or invalid. Please sign in again.' });
  }

  const email = (decoded.email || '').trim().toLowerCase();
  if (!email) {
    return res.status(401).json({ error: 'Account email is required to access workspaces.' });
  }

  try {
    req.authUser = await upsertUserRecord({
      uid: decoded.uid,
      email,
      displayName: decoded.name || '',
      photoUrl: decoded.picture || '',
    });
    next();
  } catch (err) {
    console.error('User profile sync error:', err);
    return res.status(500).json({ error: 'Could not sync your account profile. Please try again.' });
  }
};

export const requireVendorAdmin = (req: AuthRequest, res: Response, next: NextFunction) => {
  if (!req.authUser?.isVendorAdmin) {
    return res.status(403).json({ error: 'Vendor administrator access is required.' });
  }
  next();
};

export function requireOrgMember(options?: {
  allowedRoles?: OrgRole[];
  allowWhenExpired?: boolean;
}) {
  return async (req: AuthRequest, res: Response, next: NextFunction) => {
    if (!req.authUser) {
      return res.status(401).json({ error: 'Authentication required.' });
    }

    const orgId = String(req.params.orgId || '').trim();
    if (!orgId) {
      return res.status(400).json({ error: 'Organization identifier is required.' });
    }

    try {
      const rows = await db
        .select({
          org: organizations,
          member: orgMemberships,
        })
        .from(orgMemberships)
        .innerJoin(organizations, eq(orgMemberships.organizationId, organizations.id))
        .where(
          and(
            eq(orgMemberships.organizationId, orgId),
            eq(orgMemberships.userUid, req.authUser.uid)
          )
        )
        .limit(1);

      const hit = rows[0];
      if (!hit) {
        return res.status(403).json({ error: 'You do not have an assigned seat in this organization.' });
      }

      if (hit.member.seatStatus !== 'ACTIVE') {
        return res.status(403).json({
          error:
            hit.member.seatStatus === 'SUSPENDED'
              ? 'Your organization seat has been suspended by an administrator.'
              : 'Please accept your organization invitation before accessing workspace data.',
        });
      }

      const role = hit.member.role as OrgRole;
      if (options?.allowedRoles && !options.allowedRoles.includes(role)) {
        return res.status(403).json({ error: 'Your role does not permit this action.' });
      }

      const licenseExpired = hit.org.licenseExpiresAt.getTime() <= Date.now();
      if (licenseExpired && !options?.allowWhenExpired) {
        return res.status(402).json({
          error:
            'Organization license has expired. An Organization Admin must apply a renewed license key.',
          code: 'LICENSE_EXPIRED',
        });
      }

      req.orgContext = {
        organizationId: hit.org.id,
        organizationName: hit.org.name,
        orgCode: hit.org.orgCode,
        planTier: hit.org.planTier,
        seatLimit: hit.org.seatLimit,
        licenseExpiresAt: hit.org.licenseExpiresAt,
        licenseExpired,
        membershipId: hit.member.id,
        role,
        seatStatus: hit.member.seatStatus as SeatStatus,
      };

      // Update last active asynchronously without blocking request
      db.update(orgMemberships)
        .set({ lastActiveAt: new Date() })
        .where(eq(orgMemberships.id, hit.member.id))
        .catch(() => {});

      next();
    } catch (err) {
      console.error('Organization membership check failed:', err);
      return res.status(500).json({ error: 'Could not verify organization access.' });
    }
  };
}
