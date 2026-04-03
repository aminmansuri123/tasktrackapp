const SiteSettings = require('../models/SiteSettings');

async function getSiteSettings() {
  let s = await SiteSettings.findOne();
  if (!s) {
    s = await SiteSettings.create({});
  }
  if (s.registrationMode === 'email_list' || s.registrationMode === 'domain_list') {
    s.registrationMode = 'restricted';
    await s.save();
  }
  return s;
}

function normalizeEmailEntry(s) {
  return String(s || '').toLowerCase().trim();
}

function normalizeDomainEntry(s) {
  return String(s || '')
    .toLowerCase()
    .trim()
    .replace(/^@+/, '')
    .replace(/^https?:\/\//, '')
    .split('/')[0]
    .trim();
}

/**
 * @returns {Promise<string|null>} null if allowed, 'APPROVAL_REQUIRED' if neither email nor domain match
 */
async function assertRegistrationAllowed(email) {
  const s = await getSiteSettings();
  if (!s || s.registrationMode === 'open') return null;

  const em = normalizeEmailEntry(email);
  if (!em || !em.includes('@')) {
    return 'APPROVAL_REQUIRED';
  }

  const allowedEmails = new Set((s.allowedEmails || []).map(normalizeEmailEntry).filter(Boolean));
  if (allowedEmails.has(em)) return null;

  const dom = em.slice(em.indexOf('@') + 1);
  const domains = (s.allowedDomains || []).map(normalizeDomainEntry).filter(Boolean);
  const domainMatch = domains.some((d) => {
    if (!d) return false;
    if (dom === d) return true;
    return dom.endsWith(`.${d}`);
  });
  if (domainMatch) return null;

  if (allowedEmails.size === 0 && domains.length === 0) return null;

  return 'APPROVAL_REQUIRED';
}

function sanitizeBlockedLists(body) {
  const rawE = body?.blockedEmails;
  const arrE = Array.isArray(rawE) ? rawE : typeof rawE === 'string' ? rawE.split(/[\n,;]+/) : [];
  const blockedEmails = [...new Set(arrE.map(normalizeEmailEntry).filter(Boolean))];
  const rawD = body?.blockedDomains;
  const arrD = Array.isArray(rawD) ? rawD : typeof rawD === 'string' ? rawD.split(/[\n,;]+/) : [];
  const blockedDomains = [...new Set(arrD.map(normalizeDomainEntry).filter(Boolean))];
  return { blockedEmails, blockedDomains };
}

function sanitizeReportToOptions(body) {
  const raw = body?.reportToOptions;
  if (!Array.isArray(raw)) return [];
  const out = [];
  let n = 0;
  for (const o of raw) {
    if (!o || typeof o !== 'object') continue;
    const label = String(o.label || '')
      .trim()
      .slice(0, 200);
    if (!label) continue;
    let id = String(o.id || '')
      .trim()
      .slice(0, 80);
    if (!id) id = `rt_${Date.now()}_${n++}`;
    out.push({
      id,
      label,
      disabled: !!o.disabled,
    });
    if (out.length >= 200) break;
  }
  return out;
}

function sanitizePolicyBody(body) {
  const registrationMode = body?.registrationMode;
  if (!['open', 'restricted'].includes(registrationMode)) {
    return { error: 'Invalid registrationMode (must be "open" or "restricted")' };
  }
  let allowedEmails = [];
  let allowedDomains = [];
  if (registrationMode === 'restricted') {
    const rawE = body?.allowedEmails;
    const arrE = Array.isArray(rawE) ? rawE : typeof rawE === 'string' ? rawE.split(/[\n,;]+/) : [];
    allowedEmails = [...new Set(arrE.map(normalizeEmailEntry).filter(Boolean))];
    const rawD = body?.allowedDomains;
    const arrD = Array.isArray(rawD) ? rawD : typeof rawD === 'string' ? rawD.split(/[\n,;]+/) : [];
    allowedDomains = [...new Set(arrD.map(normalizeDomainEntry).filter(Boolean))];
  }
  return { registrationMode, allowedEmails, allowedDomains };
}

/**
 * @returns {Promise<string|null>} null if allowed, 'BLOCKED' if email/domain is blocklisted
 */
async function assertEmailNotBlocked(email) {
  const s = await getSiteSettings();
  const em = normalizeEmailEntry(email);
  if (!em || !em.includes('@')) return null;
  const blockedEmails = new Set((s.blockedEmails || []).map(normalizeEmailEntry).filter(Boolean));
  if (blockedEmails.has(em)) return 'BLOCKED';
  const dom = em.slice(em.indexOf('@') + 1);
  const blockedDomains = (s.blockedDomains || []).map(normalizeDomainEntry).filter(Boolean);
  for (const d of blockedDomains) {
    if (!d) continue;
    if (dom === d) return 'BLOCKED';
    if (dom.endsWith(`.${d}`)) return 'BLOCKED';
  }
  return null;
}

module.exports = {
  getSiteSettings,
  assertRegistrationAllowed,
  assertEmailNotBlocked,
  sanitizeBlockedLists,
  sanitizePolicyBody,
  sanitizeReportToOptions,
  normalizeEmailEntry,
  normalizeDomainEntry,
};
