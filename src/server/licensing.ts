import crypto from 'node:crypto';
import type { LicensePlanTier, OfflineLeaseToken, OrgRole } from '../types/enterprise.ts';

export interface SignedLicensePayload {
  id: string;
  org: string;
  tier: LicensePlanTier;
  seats: number;
  iat: number; // unix seconds
  exp: number; // unix seconds
}

interface KeyPairState {
  privateKey: crypto.KeyObject;
  publicKey: crypto.KeyObject;
  publicKeyPem: string;
  ephemeral: boolean;
}

let cachedKeyPair: KeyPairState | null = null;

function getSigningKeyPair(): KeyPairState {
  if (cachedKeyPair) return cachedKeyPair;

  const envPemRaw = process.env.LICENSE_SIGNING_PRIVATE_KEY?.trim();
  if (envPemRaw) {
    try {
      const pem = envPemRaw.includes('BEGIN PRIVATE KEY')
        ? envPemRaw
        : Buffer.from(envPemRaw, 'base64').toString('utf8');
      const privateKey = crypto.createPrivateKey(pem);
      const publicKey = crypto.createPublicKey(privateKey);
      const publicKeyPem = publicKey.export({ type: 'spki', format: 'pem' }).toString();
      cachedKeyPair = { privateKey, publicKey, publicKeyPem, ephemeral: false };
      return cachedKeyPair;
    } catch (err) {
      console.error('Invalid LICENSE_SIGNING_PRIVATE_KEY format; falling back to in-memory Ed25519 keypair:', err);
    }
  }

  const generated = crypto.generateKeyPairSync('ed25519');
  const publicKeyPem = generated.publicKey.export({ type: 'spki', format: 'pem' }).toString();
  cachedKeyPair = {
    privateKey: generated.privateKey,
    publicKey: generated.publicKey,
    publicKeyPem,
    ephemeral: true,
  };
  return cachedKeyPair;
}

export function getPublicKeyPem(): string {
  return getSigningKeyPair().publicKeyPem;
}

export function mintSignedLicenseKey(params: {
  id?: string;
  targetOrgName: string;
  planTier: LicensePlanTier;
  seatLimit: number;
  validDays: number;
}): {
  licenseId: string;
  licenseKey: string;
  issuedAt: Date;
  expiresAt: Date;
  payload: SignedLicensePayload;
} {
  const { privateKey } = getSigningKeyPair();
  const nowSec = Math.floor(Date.now() / 1000);
  const clampedSeats = Math.max(1, Math.min(1000, Math.floor(params.seatLimit)));
  const clampedDays = Math.max(7, Math.min(3650, Math.floor(params.validDays)));
  const expSec = nowSec + clampedDays * 86400;
  const licenseId =
    params.id || `LIC-${new Date().getFullYear()}-${crypto.randomBytes(4).toString('hex').toUpperCase()}`;

  const payload: SignedLicensePayload = {
    id: licenseId,
    org: params.targetOrgName.trim().slice(0, 80),
    tier: params.planTier,
    seats: clampedSeats,
    iat: nowSec,
    exp: expSec,
  };

  const payloadB64 = Buffer.from(JSON.stringify(payload), 'utf8').toString('base64url');
  const signature = crypto.sign(null, Buffer.from(payloadB64, 'utf8'), privateKey).toString('base64url');
  const licenseKey = `GGPRO-${payloadB64}.${signature}`;

  return {
    licenseId,
    licenseKey,
    issuedAt: new Date(nowSec * 1000),
    expiresAt: new Date(expSec * 1000),
    payload,
  };
}

export function parseAndVerifyLicenseKey(rawKey: string): {
  valid: boolean;
  signatureVerified: boolean;
  expired: boolean;
  payload?: SignedLicensePayload;
  error?: string;
} {
  const trimmed = rawKey.trim();
  if (!trimmed.startsWith('GGPRO-')) {
    return { valid: false, signatureVerified: false, expired: false, error: 'License key must begin with GGPRO-' };
  }
  const body = trimmed.slice('GGPRO-'.length);
  const parts = body.split('.');
  if (parts.length !== 2 || !parts[0] || !parts[1]) {
    return { valid: false, signatureVerified: false, expired: false, error: 'Malformed license key structure' };
  }

  const [payloadB64, sigB64] = parts;
  let payload: SignedLicensePayload;
  try {
    const decoded = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8'));
    if (
      !decoded ||
      typeof decoded.id !== 'string' ||
      typeof decoded.org !== 'string' ||
      typeof decoded.seats !== 'number' ||
      typeof decoded.exp !== 'number'
    ) {
      return { valid: false, signatureVerified: false, expired: false, error: 'License payload missing required fields' };
    }
    payload = decoded as SignedLicensePayload;
  } catch {
    return { valid: false, signatureVerified: false, expired: false, error: 'Could not decode license payload' };
  }

  const { publicKey } = getSigningKeyPair();
  let signatureVerified = false;
  try {
    signatureVerified = crypto.verify(
      null,
      Buffer.from(payloadB64, 'utf8'),
      publicKey,
      Buffer.from(sigB64, 'base64url')
    );
  } catch {
    signatureVerified = false;
  }

  const nowSec = Math.floor(Date.now() / 1000);
  const expired = payload.exp <= nowSec;

  return {
    valid: signatureVerified && !expired,
    signatureVerified,
    expired,
    payload,
    error: !signatureVerified
      ? 'Cryptographic signature verification failed'
      : expired
      ? 'License key has expired'
      : undefined,
  };
}

export function createSignedOfflineLease(params: {
  orgId: string;
  orgName: string;
  uid: string;
  email: string;
  role: OrgRole;
  seatLimit: number;
  planTier: LicensePlanTier;
  licenseExpiresAt: Date;
  graceDays?: number;
}): OfflineLeaseToken {
  const { privateKey } = getSigningKeyPair();
  const now = new Date();
  const graceMs = (params.graceDays ?? 7) * 86400 * 1000;
  const leaseExpiryMs = Math.min(now.getTime() + graceMs, params.licenseExpiresAt.getTime());
  const leaseBody = {
    orgId: params.orgId,
    orgName: params.orgName,
    uid: params.uid,
    email: params.email,
    role: params.role,
    seatLimit: params.seatLimit,
    planTier: params.planTier,
    issuedAt: now.toISOString(),
    expiresAt: new Date(Math.max(now.getTime(), leaseExpiryMs)).toISOString(),
  };

  const canonical = JSON.stringify(leaseBody);
  const signature = crypto.sign(null, Buffer.from(canonical, 'utf8'), privateKey).toString('base64url');
  return {
    ...leaseBody,
    signature,
  };
}
