const Workspace = require('../models/Workspace');
const { defaultWorkspaceData } = require('./defaultWorkspace');

async function ensureWorkspaceForTenantRoot(tenantRootUserId) {
  if (tenantRootUserId == null || Number.isNaN(Number(tenantRootUserId))) return null;
  const root = Number(tenantRootUserId);
  let ws = await Workspace.findOne({ tenantRootUserId: root });
  if (ws) return ws;
  try {
    return await Workspace.create({ tenantRootUserId: root, data: defaultWorkspaceData() });
  } catch (e) {
    if (e && e.code === 11000) {
      const dupKey = e.keyPattern ? Object.keys(e.keyPattern)[0] : '';
      if (dupKey === 'name') {
        console.warn('ensureWorkspaceForTenantRoot: stale name_1 unique index blocking create — dropping and retrying');
        try {
          await Workspace.collection.dropIndex('name_1');
        } catch (dropErr) {
          console.warn('Could not drop name_1 index:', dropErr.message);
        }
        try {
          return await Workspace.create({ tenantRootUserId: root, data: defaultWorkspaceData() });
        } catch (retryErr) {
          if (retryErr && retryErr.code === 11000) {
            return Workspace.findOne({ tenantRootUserId: root });
          }
          throw retryErr;
        }
      }
      return Workspace.findOne({ tenantRootUserId: root });
    }
    throw e;
  }
}

module.exports = { ensureWorkspaceForTenantRoot };
