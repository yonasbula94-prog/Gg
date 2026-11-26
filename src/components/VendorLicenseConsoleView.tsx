import React, { useEffect, useState } from 'react';
import { Ban, Copy, KeyRound, Plus, RefreshCw, ShieldCheck } from 'lucide-react';
import type { LicensePlanTier, OrgLicenseRecord } from '../types/enterprise';
import { fetchVendorLicenses, mintVendorLicense, revokeVendorLicense } from '../lib/apiClient';

interface VendorLicenseConsoleViewProps {
  onNotify: (msg: string) => void;
  onBack?: () => void;
}

export const VendorLicenseConsoleView: React.FC<VendorLicenseConsoleViewProps> = ({
  onNotify,
  onBack,
}) => {
  const [licenses, setLicenses] = useState<OrgLicenseRecord[]>([]);
  const [publicKeyPem, setPublicKeyPem] = useState('');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [targetOrgName, setTargetOrgName] = useState('');
  const [planTier, setPlanTier] = useState<LicensePlanTier>('BUSINESS');
  const [seatLimit, setSeatLimit] = useState(10);
  const [validDays, setValidDays] = useState(365);
  const [notes, setNotes] = useState('');
  const [minting, setMinting] = useState(false);
  const [lastMintedKey, setLastMintedKey] = useState<string | null>(null);

  const loadLicenses = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetchVendorLicenses();
      setLicenses(res.licenses);
      setPublicKeyPem(res.publicKeyPem);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not load vendor license ledger.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadLicenses();
  }, []);

  const handleMint = async (e: React.FormEvent) => {
    e.preventDefault();
    setMinting(true);
    setError(null);
    try {
      const res = await mintVendorLicense({
        targetOrgName: targetOrgName.trim(),
        planTier,
        seatLimit,
        validDays,
        notes: notes.trim() || undefined,
      });
      setLastMintedKey(res.license.licenseKey);
      setTargetOrgName('');
      setNotes('');
      onNotify(
        `Minted signed ${res.license.planTier} key for ${res.license.targetOrgName} (${res.license.seatLimit} seats).`
      );
      await loadLicenses();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not mint license key.');
    } finally {
      setMinting(false);
    }
  };

  const handleCopy = async (text: string, label: string) => {
    try {
      await navigator.clipboard.writeText(text);
      onNotify(`Copied ${label} to clipboard.`);
    } catch {
      onNotify('Copy failed — select the key text manually.');
    }
  };

  const handleRevoke = async (lic: OrgLicenseRecord) => {
    try {
      await revokeVendorLicense(lic.id);
      onNotify(`Revoked license key ${lic.id}.`);
      await loadLicenses();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not revoke key.');
    }
  };

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-200 rounded-xl p-6 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="font-display text-lg font-bold text-slate-900 flex items-center gap-2">
            <KeyRound className="w-5 h-5 text-[#1B365D]" />
            <span>Vendor License Key Console</span>
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            Mint cryptographically signed Ed25519 organization license keys to monetize seats per
            client company.
          </p>
        </div>
        <div className="flex items-center gap-2">
          {onBack && (
            <button
              type="button"
              onClick={onBack}
              className="px-3 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700 cursor-pointer"
            >
              Back to Workspace Setup
            </button>
          )}
          <button
            type="button"
            onClick={loadLicenses}
            className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700 cursor-pointer"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin' : ''}`} />
            <span>Refresh Ledger</span>
          </button>
        </div>
      </div>

      {error && (
        <div className="rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-xs text-red-800">
          {error}
        </div>
      )}

      {/* Mint Form */}
      <section className="bg-white border border-slate-200 rounded-xl p-6 space-y-5">
        <h3 className="font-display text-base font-bold text-slate-900">
          Generate New Signed Organization Key
        </h3>

        <form onSubmit={handleMint} className="grid grid-cols-1 md:grid-cols-12 gap-4 items-end">
          <div className="md:col-span-4">
            <label htmlFor="v-org" className="block text-xs font-medium text-slate-700 mb-1">
              Client Organization Name
            </label>
            <input
              id="v-org"
              type="text"
              required
              value={targetOrgName}
              onChange={(e) => setTargetOrgName(e.target.value)}
              placeholder="e.g. Horizon Trading LLC"
              className="w-full min-h-[40px] rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none"
            />
          </div>

          <div className="md:col-span-2">
            <label htmlFor="v-tier" className="block text-xs font-medium text-slate-700 mb-1">
              Plan Tier
            </label>
            <select
              id="v-tier"
              value={planTier}
              onChange={(e) => setPlanTier(e.target.value as LicensePlanTier)}
              className="w-full min-h-[40px] rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none"
            >
              <option value="STARTER">STARTER</option>
              <option value="BUSINESS">BUSINESS</option>
              <option value="ENTERPRISE">ENTERPRISE</option>
            </select>
          </div>

          <div className="md:col-span-2">
            <label htmlFor="v-seats" className="block text-xs font-medium text-slate-700 mb-1">
              Licensed Seats
            </label>
            <input
              id="v-seats"
              type="number"
              min={1}
              max={1000}
              required
              value={seatLimit}
              onChange={(e) => setSeatLimit(Number(e.target.value))}
              className="w-full min-h-[40px] rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none"
            />
          </div>

          <div className="md:col-span-2">
            <label htmlFor="v-days" className="block text-xs font-medium text-slate-700 mb-1">
              Validity (Days)
            </label>
            <input
              id="v-days"
              type="number"
              min={7}
              max={3650}
              required
              value={validDays}
              onChange={(e) => setValidDays(Number(e.target.value))}
              className="w-full min-h-[40px] rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none"
            />
          </div>

          <div className="md:col-span-2">
            <button
              type="submit"
              disabled={minting}
              className="w-full min-h-[40px] inline-flex items-center justify-center gap-1.5 rounded-lg bg-[#1B365D] hover:bg-[#152a48] text-white px-4 py-2 text-xs font-semibold cursor-pointer whitespace-nowrap disabled:opacity-50"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>{minting ? 'Signing...' : 'Mint Key'}</span>
            </button>
          </div>

          <div className="md:col-span-12">
            <label htmlFor="v-notes" className="block text-xs font-medium text-slate-700 mb-1">
              Invoice / Contract Reference (optional)
            </label>
            <input
              id="v-notes"
              type="text"
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              placeholder="e.g. Annual invoice #INV-2026-088 — 10 seats paid via wire"
              className="w-full min-h-[38px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none"
            />
          </div>
        </form>

        {lastMintedKey && (
          <div className="rounded-lg border border-emerald-200 bg-emerald-50 p-4 space-y-2">
            <div className="flex items-center justify-between gap-2">
              <span className="text-xs font-semibold text-emerald-900 flex items-center gap-1.5">
                <ShieldCheck className="w-4 h-4" />
                <span>Newly Signed License Key Ready for Delivery</span>
              </span>
              <button
                type="button"
                onClick={() => handleCopy(lastMintedKey, 'license key')}
                className="inline-flex items-center gap-1 px-3 py-1 rounded-md bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-semibold cursor-pointer"
              >
                <Copy className="w-3 h-3" />
                <span>Copy Key</span>
              </button>
            </div>
            <div className="font-mono text-[11px] break-all bg-white border border-emerald-200 rounded p-2.5 text-slate-900">
              {lastMintedKey}
            </div>
          </div>
        )}
      </section>

      {/* Issued Keys Ledger */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200">
          <h3 className="font-display text-base font-bold text-slate-900">
            Issued Organization License Keys ({licenses.length})
          </h3>
        </div>

        {licenses.length === 0 ? (
          <div className="p-10 text-center text-xs text-slate-500">
            No license keys minted yet. Use the form above to sign your first client license key.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-[11px] font-semibold text-slate-500">
                  <th className="py-2.5 px-4">License ID &amp; Client</th>
                  <th className="py-2.5 px-3">Plan &amp; Seats</th>
                  <th className="py-2.5 px-3">Expires</th>
                  <th className="py-2.5 px-3">Status</th>
                  <th className="py-2.5 px-4 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-xs">
                {licenses.map((lic) => (
                  <tr key={lic.id} className="hover:bg-slate-50/70">
                    <td className="py-3 px-4">
                      <div className="font-semibold text-slate-900">{lic.targetOrgName}</div>
                      <div className="text-[11px] font-mono text-slate-500">
                        {lic.id}
                        {lic.notes ? ` · ${lic.notes}` : ''}
                      </div>
                    </td>
                    <td className="py-3 px-3 font-mono tabular-nums text-slate-800">
                      {lic.planTier} · {lic.seatLimit} seats
                    </td>
                    <td className="py-3 px-3 font-mono tabular-nums text-slate-600">
                      {lic.expiresAt.slice(0, 10)}
                    </td>
                    <td className="py-3 px-3 font-semibold">
                      <span
                        className={
                          lic.status === 'ACTIVATED'
                            ? 'text-emerald-700'
                            : lic.status === 'REVOKED'
                            ? 'text-red-700'
                            : 'text-slate-700'
                        }
                      >
                        {lic.status}
                      </span>
                    </td>
                    <td className="py-3 px-4 text-right whitespace-nowrap space-x-2">
                      <button
                        type="button"
                        onClick={() => handleCopy(lic.licenseKey, lic.id)}
                        className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md border border-slate-200 bg-white hover:bg-slate-100 text-[11px] font-medium text-slate-700 cursor-pointer"
                      >
                        <Copy className="w-3 h-3" />
                        <span>Copy Key</span>
                      </button>
                      {lic.status !== 'REVOKED' && (
                        <button
                          type="button"
                          onClick={() => handleRevoke(lic)}
                          className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md border border-red-200 bg-white hover:bg-red-50 text-[11px] font-medium text-red-700 cursor-pointer"
                        >
                          <Ban className="w-3 h-3" />
                          <span>Revoke</span>
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>

      {/* Public Verification Key Reference */}
      {publicKeyPem && (
        <section className="bg-white border border-slate-200 rounded-xl p-5 space-y-2">
          <div className="flex items-center justify-between">
            <h3 className="text-xs font-semibold text-slate-800">
              Ed25519 Public Verification Key (for Tauri Offline Lease Checks)
            </h3>
            <button
              type="button"
              onClick={() => handleCopy(publicKeyPem, 'public key PEM')}
              className="text-xs font-medium text-[#1B365D] hover:underline cursor-pointer"
            >
              Copy PEM
            </button>
          </div>
          <pre className="text-[11px] font-mono bg-slate-50 border border-slate-200 rounded-lg p-3 overflow-x-auto text-slate-700">
            {publicKeyPem}
          </pre>
        </section>
      )}
    </div>
  );
};
