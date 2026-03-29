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

module.exports = {
  getSiteSettings,
  assertRegistrationAllowed,
  sanitizePolicyBody,
  normalizeEmailEntry,
  normalizeDomainEntry,
};
