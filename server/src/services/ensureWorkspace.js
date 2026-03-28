const Workspace = require('../models/Workspace');
const { defaultWorkspaceData } = require('./defaultWorkspace');

/**
 * Ensures a Workspace document exists for the tenant root. Handles E11000 races on unique tenantRootUserId.
 */
async function ensureWorkspaceForTenantRoot(tenantRootUserId) {
  if (tenantRootUserId == null || Number.isNaN(Number(tenantRootUserId))) return null;
  const root = Number(tenantRootUserId);
  let ws = await Workspace.findOne({ tenantRootUserId: root });
  if (ws) return ws;
  try {
    return await Workspace.create({ tenantRootUserId: root, data: defaultWorkspaceData() });
  } catch (e) {
    if (e && e.code === 11000) {
      return Workspace.findOne({ tenantRootUserId: root });
    }
    throw e;
  }
}

module.exports = { ensureWorkspaceForTenantRoot };
