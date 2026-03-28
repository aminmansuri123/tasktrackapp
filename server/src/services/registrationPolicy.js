const SiteSettings = require('../models/SiteSettings');

async function getSiteSettings() {
  let s = await SiteSettings.findOne();
  if (!s) {
    s = await SiteSettings.create({});
  }
  return s;
}

function normalizeEmailEntry(s) {
  return String(s || '')
    .toLowerCase()
    .trim();
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
 * @returns {Promise<string|null>} Error message to send to client, or null if allowed
 */
async function assertRegistrationAllowed(email) {
  const s = await getSiteSettings();
  if (!s || s.registrationMode === 'open') return null;

  const em = normalizeEmailEntry(email);
  if (!em || !em.includes('@')) {
    return 'Creation not allowed';
  }

  if (s.registrationMode === 'email_list') {
    const allowed = new Set((s.allowedEmails || []).map(normalizeEmailEntry).filter(Boolean));
    if (!allowed.has(em)) return 'Creation not allowed';
    return null;
  }

  if (s.registrationMode === 'domain_list') {
    const dom = em.slice(em.indexOf('@') + 1);
    const domains = (s.allowedDomains || []).map(normalizeDomainEntry).filter(Boolean);
    const ok = domains.some((d) => {
      if (!d) return false;
      if (dom === d) return true;
      return dom.endsWith(`.${d}`);
    });
    if (!ok) return 'Creation not allowed';
    return null;
  }

  return null;
}

function sanitizePolicyBody(body) {
  const registrationMode = body?.registrationMode;
  if (!['open', 'email_list', 'domain_list'].includes(registrationMode)) {
    return { error: 'Invalid registrationMode' };
  }
  let allowedEmails = [];
  let allowedDomains = [];
  if (registrationMode === 'open') {
    return { registrationMode, allowedEmails, allowedDomains };
  }
  if (registrationMode === 'email_list') {
    const raw = body?.allowedEmails;
    const arr = Array.isArray(raw) ? raw : typeof raw === 'string' ? raw.split(/[\n,;]+/) : [];
    allowedEmails = [...new Set(arr.map(normalizeEmailEntry).filter(Boolean))];
  }
  if (registrationMode === 'domain_list') {
    const raw = body?.allowedDomains;
    const arr = Array.isArray(raw) ? raw : typeof raw === 'string' ? raw.split(/[\n,;]+/) : [];
    allowedDomains = [...new Set(arr.map(normalizeDomainEntry).filter(Boolean))];
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
