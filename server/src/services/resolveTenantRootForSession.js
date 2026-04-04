const Workspace = require('../models/Workspace');

/**
 * Single source of truth for "which tenant workspace" a user belongs to — must match
 * everywhere we expose tenantRootUserId to the client or attach req.user.
 */
async function resolveTenantRootForSession(doc) {
  if (!doc) return null;
  if (doc.isMaster) return null;

  const rawTr = doc.tenantRootUserId;
  const hasTr = rawTr != null && rawTr !== '' && !Number.isNaN(Number(rawTr));
  const trNum = hasTr ? Number(rawTr) : null;
  const uidNum = Number(doc.userId);

  if (doc.role === 'admin' && !doc.isMaster) {
    if (!hasTr) {
      return uidNum;
    }
    if (trNum === uidNum) {
      return trNum;
    }
    const ownsWorkspace = await Workspace.findOne({ tenantRootUserId: doc.userId }).select('_id').lean();
    if (ownsWorkspace) {
      return uidNum;
    }
    return trNum;
  }
  if (hasTr) {
    return trNum;
  }
  return null;
}

module.exports = { resolveTenantRootForSession };
