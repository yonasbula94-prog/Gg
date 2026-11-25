import React, { useRef, useState } from 'react';
import {
  Building2,
  Check,
  ChevronDown,
  ChevronRight,
  Image as ImageIcon,
  KeyRound,
  Lock,
  MapPin,
  Palette,
  Plus,
  RotateCcw,
  Save,
  ShieldCheck,
  Trash2,
  Unlock,
  Upload,
} from 'lucide-react';
import { contrastRatio, DOCUMENT_FONT_OPTIONS, EXCEL_THEMES, SAFE_FONT_NAME } from '../constants/themes';
import {
  hasCustomAdminPassword,
  saveCompanyProfile,
  setAdminPassword,
  verifyAdminPassword,
} from '../store/storage';
import {
  CompanyProfile,
  ConfirmDialogRequest,
  DocumentFontId,
  ThemeCustomization,
  ThemePalette,
  ThemeVariantId,
} from '../types/procurement';
import { HelpTip } from './HelpTip';

interface ExcelStudioViewProps {
  themeId: ThemeVariantId;
  theme: ThemePalette;
  themeCustomization: ThemeCustomization;
  company: CompanyProfile;
  onSelectTheme: (id: ThemeVariantId) => void;
  onUpdateThemeCustomization: (patch: Partial<ThemeCustomization>) => void;
  onResetThemeCustomization: () => void;
  onUpdateCompany: (patch: Partial<CompanyProfile>) => void;
  onResetCompany: () => void;
  onRequestConfirm: (req: ConfirmDialogRequest) => void;
  onNotify: (msg: string) => void;
}

export const ExcelStudioView: React.FC<ExcelStudioViewProps> = ({
  themeId,
  theme,
  themeCustomization,
  company,
  onSelectTheme,
  onUpdateThemeCustomization,
  onResetThemeCustomization,
  onUpdateCompany,
  onResetCompany,
  onRequestConfirm,
  onNotify,
}) => {
  // Top Section: End-User Directory (name + saved address)
  const [newEndUser, setNewEndUser] = useState('');
  const [newEndUserAddress, setNewEndUserAddress] = useState('');

  // Bottom Section: Admin-Locked & Collapsed until unlocked
  const [adminUnlocked, setAdminUnlocked] = useState(false);
  const [showUnlockPrompt, setShowUnlockPrompt] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const [passwordError, setPasswordError] = useState('');
  const [showChangePass, setShowChangePass] = useState(false);
  const [newPass, setNewPass] = useState('');
  const [confirmPass, setConfirmPass] = useState('');
  const [changePassMsg, setChangePassMsg] = useState('');
  const [customPassConfigured, setCustomPassConfigured] = useState(() =>
    hasCustomAdminPassword()
  );
  const [showThemeEditor, setShowThemeEditor] = useState(false);
  const [deviceFonts, setDeviceFonts] = useState<string[]>([]);
  const [fontNote, setFontNote] = useState('');
  const canUseDeviceFonts = typeof window !== 'undefined' && 'queryLocalFonts' in window;
  const loadDeviceFonts = async () => {
    setFontNote('');
    try {
      const w = window as unknown as { queryLocalFonts: () => Promise<{ family: string }[]> };
      const fonts = await w.queryLocalFonts();
      const names = Array.from(new Set(fonts.map((f) => f.family)))
        .filter((n) => SAFE_FONT_NAME.test(n))
        .sort((a, b) => a.localeCompare(b));
      setDeviceFonts(names);
      setFontNote(`${names.length} fonts found on this device.`);
    } catch {
      setFontNote('Font access was not allowed. Allow local fonts for this site in browser settings, then try again.');
    }
  };
  const [hexDrafts, setHexDrafts] = useState<Record<string, string>>({});

  const logoInputRef = useRef<HTMLInputElement | null>(null);

  const hasCustomTheme = Boolean(
    themeCustomization.primary ||
      themeCustomization.accent ||
      themeCustomization.totalBg ||
      themeCustomization.secondaryBg ||
      themeCustomization.headerText ||
      themeCustomization.bodyText ||
      (themeCustomization.fontFamily && themeCustomization.fontFamily !== 'Calibri')
  );

  // ==================== END-USER DIRECTORY HANDLERS ====================
  const handleAddEndUser = () => {
    const trimmedName = newEndUser.trim();
    const trimmedAddr = newEndUserAddress.trim();
    if (!trimmedName) return;
    if (company.endUsers.includes(trimmedName)) {
      onNotify(`"${trimmedName}" is already in the End-User list.`);
      return;
    }
    const withoutOther = company.endUsers.filter((u) => u !== 'OTHER');
    onUpdateCompany({
      endUsers: [...withoutOther, trimmedName, 'OTHER'],
      endUserAddresses: {
        ...(company.endUserAddresses || {}),
        [trimmedName]: trimmedAddr,
      },
    });
    setNewEndUser('');
    setNewEndUserAddress('');
    onNotify(`Saved "${trimmedName}" to End-User directory.`);
  };

  const handleUpdateEndUserAddress = (userName: string, address: string) => {
    onUpdateCompany({
      endUserAddresses: {
        ...(company.endUserAddresses || {}),
        [userName]: address,
      },
    });
  };

  const handleRemoveEndUser = (name: string) => {
    if (name === 'OTHER') return;
    onRequestConfirm({
      title: `Remove "${name}"?`,
      message:
        'This removes the End User and its saved address from the PI dropdown list. Existing jobs are not affected.',
      confirmLabel: 'Remove',
      destructive: true,
      onConfirm: () => {
        const nextAddrs = { ...(company.endUserAddresses || {}) };
        delete nextAddrs[name];
        onUpdateCompany({
          endUsers: company.endUsers.filter((u) => u !== name),
          endUserAddresses: nextAddrs,
        });
        onNotify(`Removed "${name}" from End-User directory.`);
      },
    });
  };

  // ==================== ADMIN UNLOCK & BRANDING HANDLERS ====================
  const handleUnlockSubmit = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (verifyAdminPassword(passwordInput)) {
      setAdminUnlocked(true);
      setShowUnlockPrompt(false);
      setPasswordInput('');
      setPasswordError('');
      onNotify('Admin settings unlocked. Changes persist permanently across startup.');
    } else {
      setPasswordError('Incorrect admin password.');
    }
  };

  const handleSaveAndLockAdmin = () => {
    const ok = saveCompanyProfile(company);
    setAdminUnlocked(false);
    setShowChangePass(false);
    setShowThemeEditor(false);
    setNewPass('');
    setConfirmPass('');
    setChangePassMsg('');
    if (ok) {
      onNotify('Admin branding & style settings saved and locked.');
    } else {
      onNotify('Storage full — reduce logo size so settings can persist.');
    }
  };

  const handleChangeAdminPassword = (e: React.FormEvent) => {
    e.preventDefault();
    if (newPass.trim().length < 4) {
      setChangePassMsg('Password must be at least 4 characters.');
      return;
    }
    if (newPass.trim() !== confirmPass.trim()) {
      setChangePassMsg('Passwords do not match.');
      return;
    }
    const saved = setAdminPassword(newPass.trim());
    if (saved) {
      setCustomPassConfigured(true);
      setNewPass('');
      setConfirmPass('');
      setChangePassMsg('');
      setShowChangePass(false);
      onNotify('Admin password updated and saved.');
    } else {
      setChangePassMsg('Could not save password in browser storage.');
    }
  };

  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!adminUnlocked) return;
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith('image/')) {
      onNotify('Please choose a PNG or JPG logo image.');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === 'string') {
        onUpdateCompany({ logoDataUrl: reader.result });
        onNotify('Updated company logo across app, A4 print, and Excel exports.');
      }
    };
    reader.readAsDataURL(file);
    if (logoInputRef.current) logoInputRef.current.value = '';
  };

  // WCAG checks for the colours that carry text or highlights (AA: 4.5:1 text, 3:1 graphics and focus)
  const contrastChecks = (key: string): { label: string; fg: string; bg: string; min: number }[] => {
    const bestText = contrastRatio(theme.accent, '#FFFFFF') >= contrastRatio(theme.accent, '#0F172A') ? '#FFFFFF' : '#0F172A';
    switch (key) {
      case 'primary':
        return [{ label: 'Header text on banner', fg: theme.primaryText, bg: theme.primary, min: 4.5 }];
      case 'accent':
        return [
          { label: 'Button text on accent', fg: bestText, bg: theme.accent, min: 4.5 },
          { label: 'Accent on page background', fg: theme.accent, bg: theme.secondaryBg, min: 3 },
        ];
      case 'bodyText':
        return [{ label: 'Body text on white', fg: theme.bodyText, bg: '#FFFFFF', min: 4.5 }];
      default:
        return [];
    }
  };

  const colorPickers: {
    key: 'primary' | 'accent' | 'totalBg' | 'secondaryBg' | 'headerText' | 'bodyText';
    label: string;
    help: string;
    value: string;
  }[] = [
    {
      key: 'primary',
      label: 'Primary Banner & Header Fill',
      help: 'Main banner fill, table header background, and primary action buttons across A4 print and Excel.',
      value: theme.primary,
    },
    {
      key: 'headerText',
      label: 'Header Font Color',
      help: 'Font color for text inside the top document banner and table header row.',
      value: theme.primaryText,
    },
    {
      key: 'bodyText',
      label: 'Document Font Color',
      help: 'Font color for body text, client details, and line-item values in A4 Print Preview and Excel exports.',
      value: theme.bodyText,
    },
    {
      key: 'accent',
      label: 'Accent Rule & Highlight',
      help: 'Divider accent border under document header and highlighted costing column.',
      value: theme.accent,
    },
    {
      key: 'totalBg',
      label: 'Grand Total Row Fill',
      help: 'Background tint of the Grand Total box in A4 print and Excel total rows.',
      value: theme.totalBg,
    },
    {
      key: 'secondaryBg',
      label: 'Sub-Header & Info Box Tint',
      help: 'Light tinted background used for sub-headers and contact strip.',
      value: theme.secondaryBg,
    },
  ];

  const namedEndUsers = company.endUsers.filter((u) => u !== 'OTHER');

  return (
    <div className="space-y-6">
      {/* ==================== TOP SECTION: END-USER DIRECTORY (SEPARATED & UNLOCKED) ==================== */}
      <div className="rounded-lg border border-slate-200 bg-white p-6 space-y-5">
        <div className="flex flex-wrap items-center justify-between gap-2 pb-4 border-b border-slate-100">
          <div className="flex items-center gap-2">
            <Building2 className="w-4 h-4 text-slate-600 shrink-0" />
            <h2 className="text-base font-semibold text-slate-900">End-User Directory</h2>
            <HelpTip text="Manage your End-User dropdown list and saved addresses. When you select an End User on Proforma Invoice (PI), their saved address fills automatically while ATTN stays blank for each order." />
          </div>
          <span className="text-xs font-mono-tabular text-slate-500">
            {namedEndUsers.length} Saved Clients (+ OTHER)
          </span>
        </div>

        {/* Add New End-User + Address Form */}
        <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-end">
          <div className="md:col-span-5">
            <label
              htmlFor="new-end-user-name"
              className="block text-xs font-medium text-slate-600 mb-1"
            >
              End-User Name
            </label>
            <input
              id="new-end-user-name"
              type="text"
              value={newEndUser}
              onChange={(e) => setNewEndUser(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleAddEndUser();
              }}
              placeholder="e.g. Red Sea Maritime Corp"
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
            />
          </div>
          <div className="md:col-span-5">
            <label
              htmlFor="new-end-user-address"
              className="block text-xs font-medium text-slate-600 mb-1"
            >
              Saved Address
            </label>
            <input
              id="new-end-user-address"
              type="text"
              value={newEndUserAddress}
              onChange={(e) => setNewEndUserAddress(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleAddEndUser();
              }}
              placeholder="e.g. Harbor Rd, Massawa, Eritrea"
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
            />
          </div>
          <div className="md:col-span-2">
            <button
              type="button"
              disabled={!newEndUser.trim()}
              onClick={handleAddEndUser}
              className="w-full inline-flex items-center justify-center gap-1.5 rounded-lg px-3.5 py-2 text-xs font-semibold hover:opacity-90 cursor-pointer min-h-[40px] whitespace-nowrap disabled:opacity-40 disabled:cursor-not-allowed"
              style={{ backgroundColor: theme.primary, color: theme.primaryText }}
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Add End User</span>
            </button>
          </div>
        </div>

        {/* Saved End-Users with Inline Editable Addresses */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
          {namedEndUsers.map((user) => (
            <div
              key={user}
              className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2"
            >
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs font-semibold text-slate-900 truncate">{user}</span>
                <button
                  type="button"
                  onClick={() => handleRemoveEndUser(user)}
                  title={`Remove ${user}`}
                  aria-label={`Remove ${user}`}
                  className="p-1 rounded text-slate-400 hover:text-red-600 hover:bg-red-50 cursor-pointer shrink-0"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </div>
              <div className="flex items-center gap-2">
                <MapPin className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                <input
                  type="text"
                  aria-label={`Address for ${user}`}
                  value={company.endUserAddresses?.[user] || ''}
                  onChange={(e) => handleUpdateEndUserAddress(user, e.target.value)}
                  placeholder="Enter saved address (auto-fills on PI)..."
                  className="w-full min-h-[34px] rounded-md border border-slate-200 bg-white px-2.5 py-1 text-xs text-slate-800 focus:border-slate-900 focus:outline-none"
                />
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* ==================== BOTTOM SECTION: ADMIN-LOCKED & COLLAPSED UNTIL UNLOCKED ==================== */}
      <div className="rounded-lg border border-slate-200 bg-white overflow-visible">
        {/* Admin Bar Header (Always Visible — Body Completely Collapsed Until Unlocked) */}
        <div className="flex flex-wrap items-center justify-between gap-3 px-6 py-4">
          <div className="flex items-center gap-2.5 min-w-0">
            {adminUnlocked ? (
              <ShieldCheck className="w-4 h-4 text-emerald-700 shrink-0" />
            ) : (
              <Lock className="w-4 h-4 text-slate-500 shrink-0" />
            )}
            <h2 className="text-base font-semibold text-slate-900">
              Company Branding, Logo &amp; Document Styles (Admin)
            </h2>
            <HelpTip text="Protected by Admin Password and kept collapsed until unlocked. Contains Company Branding/Logo details and Excel Document Style Variants (v1 · v2 · v3), Colors & Fonts." />
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {!adminUnlocked ? (
              <>
                <span className="text-xs font-mono-tabular text-slate-500 hidden sm:inline">
                  Locked · Collapsed
                </span>
                <button
                  type="button"
                  onClick={() => {
                    setShowUnlockPrompt((v) => !v);
                    setPasswordError('');
                  }}
                  className="inline-flex items-center gap-1.5 rounded-lg px-3.5 py-1.5 text-xs font-semibold hover:opacity-90 cursor-pointer min-h-[36px] whitespace-nowrap"
                  style={{ backgroundColor: theme.primary, color: theme.primaryText }}
                >
                  <Unlock className="w-3.5 h-3.5" />
                  <span>Unlock Admin Settings</span>
                </button>
              </>
            ) : (
              <>
                <button
                  type="button"
                  onClick={() => {
                    setShowChangePass((v) => !v);
                    setChangePassMsg('');
                  }}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[36px] whitespace-nowrap"
                >
                  <KeyRound className="w-3.5 h-3.5" />
                  <span>Change Password</span>
                </button>
                <button
                  type="button"
                  onClick={() =>
                    onRequestConfirm({
                      title: 'Reset Company Branding & Defaults?',
                      message:
                        'This will restore the default Grand Group FZC logo, SAIF Zone address, sales@grandgroupfzc.com email, bank details, and default End-User addresses.',
                      confirmLabel: 'Reset Defaults',
                      destructive: true,
                      onConfirm: onResetCompany,
                    })
                  }
                  className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 hover:text-slate-900 cursor-pointer min-h-[36px] whitespace-nowrap"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  <span>Reset Defaults</span>
                </button>
                <button
                  type="button"
                  onClick={handleSaveAndLockAdmin}
                  className="inline-flex items-center gap-1.5 rounded-lg bg-emerald-700 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-emerald-800 cursor-pointer min-h-[36px] whitespace-nowrap"
                >
                  <Save className="w-3.5 h-3.5" />
                  <span>Save &amp; Lock</span>
                </button>
              </>
            )}
          </div>
        </div>

        {/* Inline Password Unlock Bar (Shown Only When Clicking Unlock) */}
        {!adminUnlocked && showUnlockPrompt && (
          <div className="px-6 pb-5">
            <form
              onSubmit={handleUnlockSubmit}
              className="rounded-lg border border-slate-300 bg-slate-50 p-4 flex flex-wrap items-center justify-between gap-3"
            >
              <div className="flex flex-wrap items-center gap-2.5 flex-1 min-w-[240px]">
                <label
                  htmlFor="admin-password-input"
                  className="text-xs font-semibold text-slate-800 whitespace-nowrap"
                >
                  Admin Password:
                </label>
                <input
                  id="admin-password-input"
                  type="password"
                  autoFocus
                  value={passwordInput}
                  onChange={(e) => {
                    setPasswordInput(e.target.value);
                    setPasswordError('');
                  }}
                  placeholder="Enter admin password..."
                  className="min-h-[36px] rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs text-slate-900 focus:border-slate-900 focus:outline-none w-52"
                />
                <button
                  type="submit"
                  className="inline-flex items-center gap-1.5 rounded-lg bg-slate-900 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-slate-800 cursor-pointer min-h-[36px]"
                >
                  <Unlock className="w-3.5 h-3.5" />
                  <span>Unlock</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setShowUnlockPrompt(false);
                    setPasswordInput('');
                    setPasswordError('');
                  }}
                  className="rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-100 cursor-pointer min-h-[36px]"
                >
                  Cancel
                </button>
              </div>
              <div className="text-xs">
                {passwordError ? (
                  <span className="font-semibold text-red-700">{passwordError}</span>
                ) : !customPassConfigured ? (
                  <span className="text-slate-500">
                    Initial default password:{' '}
                    <code className="font-mono-tabular font-semibold text-slate-700">
                      GGAdmin2026
                    </code>
                  </span>
                ) : null}
              </div>
            </form>
          </div>
        )}

        {/* ==================== EXPANDED ADMIN CONTENT (ONLY RENDERED ONCE UNLOCKED) ==================== */}
        {adminUnlocked && (
          <div className="px-6 pb-6 pt-5 border-t border-slate-200 space-y-8">
            {/* Optional Change Admin Password Form */}
            {showChangePass && (
              <form
                onSubmit={handleChangeAdminPassword}
                className="rounded-lg border border-slate-200 bg-slate-50 p-4 flex flex-wrap items-center gap-2.5"
              >
                <span className="text-xs font-semibold text-slate-800 whitespace-nowrap">
                  Set New Admin Password:
                </span>
                <input
                  type="password"
                  value={newPass}
                  onChange={(e) => {
                    setNewPass(e.target.value);
                    setChangePassMsg('');
                  }}
                  placeholder="New password (min 4 chars)"
                  className="min-h-[36px] rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs text-slate-900 focus:border-slate-900 focus:outline-none w-48"
                />
                <input
                  type="password"
                  value={confirmPass}
                  onChange={(e) => {
                    setConfirmPass(e.target.value);
                    setChangePassMsg('');
                  }}
                  placeholder="Confirm new password"
                  className="min-h-[36px] rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs text-slate-900 focus:border-slate-900 focus:outline-none w-48"
                />
                <button
                  type="submit"
                  className="inline-flex items-center gap-1.5 rounded-lg bg-slate-900 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-slate-800 cursor-pointer min-h-[36px]"
                >
                  Save Password
                </button>
                {changePassMsg && (
                  <span className="text-xs font-medium text-red-700">{changePassMsg}</span>
                )}
              </form>
            )}

            {/* Sub-Section 1: Company Branding & Logo */}
            <div className="space-y-4">
              <div className="flex items-center gap-1.5">
                <h3 className="text-sm font-semibold text-slate-900">
                  Company Branding &amp; Logo
                </h3>
                <HelpTip text="Updates company letterhead, contact details, TRN, bank wire instructions, and logo across all A4 prints and Excel exports." />
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                {/* Logo Preview & Upload */}
                <div className="lg:col-span-3 flex flex-col items-center justify-center rounded-lg border border-slate-200 bg-slate-50/50 p-5 text-center">
                  <div className="w-24 h-24 rounded-lg border border-slate-200 bg-white p-2 flex items-center justify-center">
                    {company.logoDataUrl ? (
                      <img
                        src={company.logoDataUrl}
                        alt={company.name}
                        className="max-w-full max-h-full object-contain"
                      />
                    ) : (
                      <ImageIcon className="w-8 h-8 text-slate-300" />
                    )}
                  </div>
                  <input
                    ref={logoInputRef}
                    type="file"
                    accept="image/png,image/jpeg"
                    onChange={handleLogoUpload}
                    className="hidden"
                  />
                  <button
                    type="button"
                    onClick={() => logoInputRef.current?.click()}
                    className="mt-3 inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[36px]"
                  >
                    <Upload className="w-3.5 h-3.5" />
                    <span>Replace Logo</span>
                  </button>
                  <span className="mt-1.5 text-[11px] text-slate-400">
                    PNG or JPG · Embedded in .xlsx
                  </span>
                </div>

                {/* Company Text Fields */}
                <div className="lg:col-span-9 grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Company Name
                    </label>
                    <input
                      type="text"
                      value={company.name}
                      onChange={(e) => onUpdateCompany({ name: e.target.value })}
                      className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Tagline
                    </label>
                    <input
                      type="text"
                      value={company.tagline}
                      onChange={(e) => onUpdateCompany({ tagline: e.target.value })}
                      className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Address Line 1
                    </label>
                    <input
                      type="text"
                      value={company.addressLine1}
                      onChange={(e) => onUpdateCompany({ addressLine1: e.target.value })}
                      className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Address Line 2
                    </label>
                    <input
                      type="text"
                      value={company.addressLine2}
                      onChange={(e) => onUpdateCompany({ addressLine2: e.target.value })}
                      className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Email</label>
                    <input
                      type="text"
                      value={company.email}
                      onChange={(e) => onUpdateCompany({ email: e.target.value })}
                      className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Phone</label>
                    <input
                      type="text"
                      value={company.phone}
                      onChange={(e) => onUpdateCompany({ phone: e.target.value })}
                      className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-slate-900 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      TRN / Tax Registration
                    </label>
                    <input
                      type="text"
                      value={company.trn}
                      onChange={(e) => onUpdateCompany({ trn: e.target.value })}
                      className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-slate-900 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Bank &amp; Wire Details (PI &amp; CI)
                    </label>
                    <input
                      type="text"
                      value={company.bankDetails}
                      onChange={(e) => onUpdateCompany({ bankDetails: e.target.value })}
                      className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-slate-900 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                </div>
              </div>
            </div>

            {/* Sub-Section 2: Excel Document Style Variants (v1 · v2 · v3), Colors & Fonts */}
            <div className="pt-6 border-t border-slate-200 space-y-5">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <h3 className="text-sm font-semibold text-slate-900">
                    Excel Document Style Variants (v1 · v2 · v3)
                  </h3>
                  <HelpTip text="Choose a preset palette or open Customize Colors & Font to edit banner, accent, total fill, sub-header, header font color, document font color, and font family across A4 Print, PDF, and Excel exports." />
                </div>
                <div className="flex items-center gap-2.5">
                  <span className="text-xs font-mono-tabular font-semibold text-[#0b54cb] hidden sm:inline">
                    {EXCEL_THEMES[themeId].shortLabel} · {theme.fontName}
                    {hasCustomTheme ? ' (Custom)' : ''}
                  </span>
                  <button
                    type="button"
                    aria-expanded={showThemeEditor}
                    onClick={() => setShowThemeEditor((v) => !v)}
                    className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-semibold text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[36px] whitespace-nowrap"
                  >
                    <Palette className="w-3.5 h-3.5" />
                    <span>Customize Colors &amp; Font</span>
                    {showThemeEditor ? (
                      <ChevronDown className="w-3.5 h-3.5 text-slate-400" />
                    ) : (
                      <ChevronRight className="w-3.5 h-3.5 text-slate-400" />
                    )}
                  </button>
                </div>
              </div>

              {/* Preset Cards */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
                {(['v1', 'v3'] as ThemeVariantId[]).map((id) => {
                  const isSelected = themeId === id;
                  const t = isSelected ? theme : EXCEL_THEMES[id];

                  return (
                    <button
                      key={id}
                      type="button"
                      onClick={() => onSelectTheme(id)}
                      className={`text-left rounded-lg border p-4 transition-colors cursor-pointer ${
                        isSelected
                          ? 'border-slate-900 ring-1 ring-slate-900 bg-slate-50/40'
                          : 'border-slate-200 hover:border-slate-300 bg-white'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-sm font-semibold text-slate-900">
                          {EXCEL_THEMES[id].name}
                        </span>
                        {isSelected && (
                          <span className="inline-flex items-center gap-1 text-xs font-semibold text-slate-900">
                            <Check className="w-3.5 h-3.5" /> Active
                          </span>
                        )}
                      </div>
                      <p className="text-xs text-slate-500 mt-1">{EXCEL_THEMES[id].subtitle}</p>

                      {/* Live Mini Excel Sheet Specimen */}
                      <div
                        className="mt-4 rounded border overflow-hidden text-[11px]"
                        style={{
                          borderColor: t.border,
                          fontFamily: t.fontCss,
                          color: t.bodyText,
                        }}
                      >
                        <div
                          className="px-2.5 py-1.5 flex items-center justify-between font-semibold"
                          style={{ backgroundColor: t.primary, color: t.primaryText }}
                        >
                          <span className="flex items-center gap-1.5 min-w-0">
                            {company.logoDataUrl && (
                              <span className="w-4 h-4 rounded-sm bg-white p-0.5 flex items-center justify-center shrink-0">
                                <img
                                  src={company.logoDataUrl}
                                  alt=""
                                  className="max-w-full max-h-full object-contain"
                                />
                              </span>
                            )}
                            <span className="truncate">{company.name}</span>
                          </span>
                          <span className="shrink-0">PROFORMA</span>
                        </div>
                        <div
                          className="px-2.5 py-1 text-[10px] border-b-2 flex justify-between"
                          style={{
                            backgroundColor: t.secondaryBg,
                            borderColor: t.accent,
                            color: t.mutedText,
                          }}
                        >
                          <span>AIAA · CIF Massawa</span>
                          <span>{t.fontName}</span>
                        </div>
                        <div className="bg-white divide-y" style={{ borderColor: t.border }}>
                          <div className="px-2.5 py-1 flex justify-between">
                            <span className="truncate pr-2">LED Runway Edge Light</span>
                            <span className="font-mono-tabular">7,866.00</span>
                          </div>
                          <div
                            className="px-2.5 py-1 flex justify-between"
                            style={{ backgroundColor: t.zebraRow }}
                          >
                            <span className="truncate pr-2">Isolation Transformer 100W</span>
                            <span className="font-mono-tabular">1,987.20</span>
                          </div>
                          <div
                            className="px-2.5 py-1.5 flex justify-between font-bold"
                            style={{ backgroundColor: t.totalBg, color: t.primary }}
                          >
                            <span>GRAND TOTAL</span>
                            <span className="font-mono-tabular">10,345.86</span>
                          </div>
                        </div>
                      </div>
                    </button>
                  );
                })}
              </div>

              {/* Collapsible Template Colors, Font Color & Font Family Customizer */}
              {showThemeEditor && (
                <div className="pt-5 border-t border-slate-200 space-y-5">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <h4 className="text-xs font-semibold text-slate-900 uppercase tracking-wide">
                        Customize Colors, Font Color &amp; Font Family
                      </h4>
                      <HelpTip text="All custom colors, font colors, and font family choices persist automatically across startup and apply to every A4 Print Preview, PDF, and exported .xlsx workbook." />
                    </div>
                    <button
                      type="button"
                      disabled={!hasCustomTheme}
                      onClick={() => {
                        onResetThemeCustomization();
                        onNotify('Reset template colors and font to preset defaults.');
                      }}
                      className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[34px] disabled:opacity-40 disabled:cursor-not-allowed"
                    >
                      <RotateCcw className="w-3.5 h-3.5" />
                      <span>Reset Colors &amp; Font</span>
                    </button>
                  </div>

                  {/* Document Font Family Row */}
                  <div className="max-w-sm space-y-1.5">
                    <label
                      htmlFor="doc-font-select"
                      className="flex items-center gap-1 text-xs font-medium text-slate-700"
                    >
                      <span>Document Font Family</span>
                      <HelpTip text="Applies to all A4 document prints and every cell inside exported Excel (.xlsx) workbooks." />
                    </label>
                    <select
                      id="doc-font-select"
                      value={theme.fontName}
                      onChange={(e) =>
                        onUpdateThemeCustomization({
                          fontFamily: e.target.value as DocumentFontId,
                        })
                      }
                      className="w-full min-h-[42px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-900 focus:border-slate-900 focus:outline-none"
                    >
                      {DOCUMENT_FONT_OPTIONS.map((f) => (
                        <option key={f.id} value={f.id}>
                          {f.label}
                        </option>
                      ))}
                      {deviceFonts.filter((n) => !DOCUMENT_FONT_OPTIONS.some((p) => p.id === n)).length > 0 && (
                        <optgroup label="Installed on this device">
                          {deviceFonts
                            .filter((n) => !DOCUMENT_FONT_OPTIONS.some((p) => p.id === n))
                            .map((n) => (
                              <option key={n} value={n}>
                                {n}
                              </option>
                            ))}
                        </optgroup>
                      )}
                      {!DOCUMENT_FONT_OPTIONS.some((p) => p.id === theme.fontName) &&
                        !deviceFonts.includes(theme.fontName) && (
                          <option value={theme.fontName}>{theme.fontName}</option>
                        )}
                    </select>
                    <div className="mt-2 flex flex-wrap items-center gap-2">
                      {canUseDeviceFonts ? (
                        <button
                          type="button"
                          onClick={loadDeviceFonts}
                          className="rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[36px]"
                        >
                          Use fonts installed on this device
                        </button>
                      ) : (
                        <span className="text-[11px] text-slate-500">
                          Device font list needs Chrome or Edge on Windows. Other devices use the presets above.
                        </span>
                      )}
                      {fontNote && <span className="text-[11px] text-slate-600">{fontNote}</span>}
                    </div>
                  </div>

                  {/* 6 Color Pickers (Includes Header Font Color & Document Font Color) */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                    {colorPickers.map((cp) => (
                      <div key={cp.key} className="space-y-1.5">
                        <label className="flex items-center gap-1 text-xs font-medium text-slate-700">
                          <span className="truncate">{cp.label}</span>
                          <HelpTip text={cp.help} />
                        </label>
                        {contrastChecks(cp.key).map((check) => {
                          const ratio = contrastRatio(check.fg, check.bg);
                          const ok = ratio >= check.min;
                          return (
                            <p
                              key={check.label}
                              className={`text-[11px] ${ok ? 'text-slate-600' : 'text-amber-800 font-medium'}`}
                            >
                              {ok ? '✓' : '⚠'} {check.label}: {ratio.toFixed(1)}:1
                              {ok ? '' : ` (needs ${check.min}:1; choose a ${check.fg === theme.accent || check.bg === theme.accent ? 'stronger' : 'darker or lighter'} shade)`}
                            </p>
                          );
                        })}
                        <div className="flex items-center gap-2 rounded-lg border border-slate-200 bg-slate-50/60 px-2.5 py-1.5 min-h-[42px]">
                          <input
                            type="color"
                            aria-label={cp.label}
                            value={cp.value}
                            onChange={(e) => {
                              const next = e.target.value.toUpperCase();
                              setHexDrafts((prev) => {
                                const copy = { ...prev };
                                delete copy[cp.key];
                                return copy;
                              });
                              onUpdateThemeCustomization({ [cp.key]: next });
                            }}
                            className="w-7 h-7 rounded border border-slate-300 bg-white cursor-pointer shrink-0 p-0.5"
                          />
                          <input
                            type="text"
                            aria-label={`${cp.label} Hex`}
                            value={hexDrafts[cp.key] ?? cp.value}
                            maxLength={7}
                            onChange={(e) => {
                              const typed = e.target.value.toUpperCase();
                              setHexDrafts((prev) => ({ ...prev, [cp.key]: typed }));
                              const normalized = typed.startsWith('#') ? typed : `#${typed}`;
                              if (/^#[0-9A-F]{6}$/.test(normalized)) {
                                onUpdateThemeCustomization({ [cp.key]: normalized });
                              }
                            }}
                            onBlur={() => {
                              setHexDrafts((prev) => {
                                const copy = { ...prev };
                                delete copy[cp.key];
                                return copy;
                              });
                            }}
                            className="w-full bg-transparent text-xs font-mono-tabular font-semibold text-slate-800 focus:outline-none uppercase"
                          />
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
