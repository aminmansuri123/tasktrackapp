const User = require('../models/User');

/** Tenant root userId for an org (from any admin user row in that org). */
function tenantRootFromUserDoc(u) {
  if (!u) return null;
  return u.tenantRootUserId != null ? u.tenantRootUserId : u.userId;
}

/**
 * Resolve tenant root from a picker id (organization primary admin or any admin in the org).
 * @returns {Promise<number|null>}
 */
async function resolveTenantRootFromAdminPicker(pickerUserId) {
  const id = typeof pickerUserId === 'number' ? pickerUserId : parseInt(String(pickerUserId), 10);
  if (Number.isNaN(id)) return null;
  const u = await User.findOne({ userId: id, isMaster: { $ne: true }, role: 'admin', isActive: true });
  if (!u) return null;
  return tenantRootFromUserDoc(u);
}

module.exports = { tenantRootFromUserDoc, resolveTenantRootFromAdminPicker };
