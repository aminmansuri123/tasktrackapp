const User = require('../models/User');

/** Tenant root userId for an org (from any user row in that org). */
function tenantRootFromUserDoc(u) {
  if (!u) return null;
  const raw = u.tenantRootUserId != null && u.tenantRootUserId !== '' ? u.tenantRootUserId : u.userId;
  if (raw == null || raw === '') return null;
  const n = Number(raw);
  return Number.isNaN(n) ? null : n;
}

/**
 * Resolve tenant root from a picker id (any active member of an org — not only role admin).
 * Registration lists account admins, but linking must work if the chosen id is a valid tenant
 * member (e.g. delegated admin stored as user, or owner row edge cases); team users share the
 * same tenantRootUserId as their admin.
 * @returns {Promise<number|null>}
 */
async function resolveTenantRootFromAdminPicker(pickerUserId) {
  const id = typeof pickerUserId === 'number' ? pickerUserId : parseInt(String(pickerUserId), 10);
  if (Number.isNaN(id)) return null;
  const u = await User.findOne({ userId: id, isMaster: { $ne: true }, isActive: true });
  if (!u) return null;
  return tenantRootFromUserDoc(u);
}

/**
 * Workspace / GridFS tenant key for the signed-in user. Fixes org admins whose User row has a
 * stale or wrong tenantRootUserId while team members still point at this admin's userId.
 * Delegated admins (userId !== tenantRootUserId, members use the org root id) are unchanged.
 */
async function resolveWorkspaceTenantRootUserId(doc) {
  if (!doc || doc.isMaster) return null;
  if (doc.role !== 'admin') {
    return tenantRootFromUserDoc(doc);
  }
  const selfId = Number(doc.userId);
  if (!Number.isFinite(selfId)) {
    return tenantRootFromUserDoc(doc);
  }
  const trRaw = doc.tenantRootUserId;
  const tr =
    trRaw != null && trRaw !== '' && !Number.isNaN(Number(trRaw)) ? Number(trRaw) : null;
  const mismatched = tr != null && Number.isFinite(tr) && tr !== selfId;
  if (mismatched) {
    const membersUnderSelf = await User.countDocuments({
      isMaster: { $ne: true },
      tenantRootUserId: selfId,
    });
    if (membersUnderSelf > 0) {
      return selfId;
    }
  }
  return tenantRootFromUserDoc(doc);
}

module.exports = {
  tenantRootFromUserDoc,
  resolveTenantRootFromAdminPicker,
  resolveWorkspaceTenantRootUserId,
};
