const express = require('express');
const User = require('../models/User');
const Workspace = require('../models/Workspace');
const { authMiddleware } = require('../middleware/auth');
const { defaultWorkspaceData, normalizeWorkspacePayload } = require('../services/defaultWorkspace');
const { ensureWorkspaceForTenantRoot } = require('../services/ensureWorkspace');
const {
  syncUsersFromClientPayload,
  deleteUsersNotInPayload,
  usersToClientShapeForTenant,
  usersToClientShapeAll,
  previousWorkspaceUserIdSet,
  mergeIncomingUsersWithDbTenantRoster,
} = require('../services/userSync');

const router = express.Router();

const EXPORT_VERSION = '15.1.0';

/**
 * Determine workspace tenant root for the current request.
 * For org-owner admins: tenantRootUserId === userId → use it.
 * For delegated admins: tenantRootUserId !== userId → use tenantRootUserId.
 * For regular users: use tenantRootUserId.
 * Fallback: use userId.
 */
function resolveTenantRoot(req) {
  if (req.user.isMaster) return null;
  const tr = req.user.tenantRootUserId;
  if (tr != null && !Number.isNaN(Number(tr))) {
    return Number(tr);
  }
  return Number(req.user.userId);
}

async function loadTenantUsers(tenantRoot, currentUserId) {
  let users = [];
  try {
    users = await usersToClientShapeForTenant(tenantRoot);
  } catch (e) {
    console.error('loadTenantUsers: usersToClientShapeForTenant failed:', e.message);
  }

  const uid = Number(currentUserId);
  if (!Number.isNaN(uid) && !users.some((u) => Number(u.id) === uid)) {
    try {
      const doc = await User.findOne({ userId: uid }).lean();
      if (doc && !doc.isMaster) {
        users.push({
          id: doc.userId,
          email: doc.email,
          name: doc.name,
          role: doc.role,
          is_active: doc.isActive,
          isMaster: doc.isMaster,
          enabledFeatures: Array.isArray(doc.enabledFeatures) ? doc.enabledFeatures : [],
        });
      }
    } catch (e) {
      console.error('loadTenantUsers: fallback user lookup failed:', e.message);
    }
  }
  return users;
}

async function loadSharedTasks(userId) {
  try {
    const userDoc = await User.findOne({ userId }).lean();
    if (!userDoc || !Array.isArray(userDoc.sharedWithTenants) || userDoc.sharedWithTenants.length === 0) {
      return [];
    }
    const sharedTasks = [];
    for (const tenantId of userDoc.sharedWithTenants) {
      const ws = await Workspace.findOne({ tenantRootUserId: tenantId }).lean();
      if (!ws || !ws.data || !Array.isArray(ws.data.tasks)) continue;
      const admin = await User.findOne({
        $or: [{ userId: tenantId }, { tenantRootUserId: tenantId, role: 'admin' }],
        isMaster: { $ne: true },
      }).lean();
      const adminName = admin ? (admin.name || admin.email) : `Org ${tenantId}`;
      const tasksForUser = ws.data.tasks.filter(
        (t) => Number(t.assigned_to) === userId
      );
      for (const t of tasksForUser) {
        sharedTasks.push({
          ...t,
          _sharedTask: true,
          _sourceWorkspace: tenantId,
          _assignedByAdmin: adminName,
        });
      }
    }
    return sharedTasks;
  } catch (e) {
    console.error('loadSharedTasks error:', e.message);
    return [];
  }
}

router.get('/debug-tenant', authMiddleware, async (req, res) => {
  try {
    const tenantRoot = resolveTenantRoot(req);
    const uid = req.user.userId;

    const adminDoc = await User.findOne({ userId: uid }).lean();
    const wsDoc = await Workspace.findOne({ tenantRootUserId: tenantRoot }).lean();
    const allTenantUsers = await User.find({
      isMaster: { $ne: true },
      $or: [{ tenantRootUserId: tenantRoot }, { userId: tenantRoot }],
    }).lean();

    const allWorkspaces = await Workspace.find().select('tenantRootUserId').lean();

    return res.json({
      resolvedTenantRoot: tenantRoot,
      reqUser: req.user,
      adminDoc: adminDoc
        ? { userId: adminDoc.userId, email: adminDoc.email, role: adminDoc.role, tenantRootUserId: adminDoc.tenantRootUserId, isMaster: adminDoc.isMaster, isActive: adminDoc.isActive }
        : null,
      workspaceExists: !!wsDoc,
      workspaceTenantRootUserId: wsDoc ? wsDoc.tenantRootUserId : null,
      workspaceDataUsersCount: wsDoc && wsDoc.data ? (wsDoc.data.users || []).length : 0,
      tenantUsersFromDb: allTenantUsers.map((u) => ({ userId: u.userId, email: u.email, tenantRootUserId: u.tenantRootUserId, role: u.role })),
      allWorkspaceTenantRoots: allWorkspaces.map((w) => w.tenantRootUserId),
    });
  } catch (e) {
    console.error('debug-tenant error:', e);
    return res.status(500).json({ error: e.message });
  }
});

router.get('/backup', authMiddleware, async (req, res) => {
  try {
    if (req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin only' });
    }
    if (req.user.isMaster) {
      return res.status(403).json({ error: 'Use a tenant admin account for workspace backup' });
    }
    const tenantRoot = resolveTenantRoot(req);
    const ws = await ensureWorkspaceForTenantRoot(tenantRoot);
    if (!ws) {
      return res.status(404).json({ error: 'No workspace' });
    }
    const normalized = normalizeWorkspacePayload(ws.data);
    normalized.users = await loadTenantUsers(tenantRoot, req.user.userId);
    return res.json({
      version: EXPORT_VERSION,
      exportDate: new Date().toISOString(),
      applicationName: 'Task Management System',
      data: normalized,
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Backup failed' });
  }
});

router.post('/restore', authMiddleware, async (req, res) => {
  try {
    if (req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin only' });
    }
    if (req.user.isMaster) {
      return res.status(403).json({ error: 'Use a tenant admin account for restore' });
    }
    const tenantRoot = resolveTenantRoot(req);
    const body = req.body || {};
    const imported = body.data || body;
    if (!imported || typeof imported !== 'object') {
      return res.status(400).json({ error: 'Invalid backup payload' });
    }
    const complete = normalizeWorkspacePayload(imported);
    let ws = await Workspace.findOne({ tenantRootUserId: tenantRoot });
    if (!ws) {
      try {
        ws = await Workspace.create({ tenantRootUserId: tenantRoot, data: complete });
      } catch (e) {
        if (e && e.code === 11000) {
          ws = await Workspace.findOne({ tenantRootUserId: tenantRoot });
        } else throw e;
      }
    }
    if (!ws) {
      return res.status(500).json({ error: 'Restore failed' });
    }
    ws.data = complete;
    ws.markModified('data');
    await ws.save();
    try {
      await syncUsersFromClientPayload(complete.users, { isAdmin: true, tenantRootUserId: tenantRoot });
      if (Array.isArray(complete.users) && complete.users.length > 0) {
        await deleteUsersNotInPayload(complete.users.map((u) => u.id), tenantRoot, null, true);
      }
    } catch (syncErr) {
      console.error('Restore: user sync error:', syncErr.message);
    }
    const normalized = normalizeWorkspacePayload(ws.data);
    normalized.users = await loadTenantUsers(tenantRoot, req.user.userId);
    return res.json(normalized);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Restore failed' });
  }
});

router.get('/', authMiddleware, async (req, res) => {
  try {
    if (req.user.isMaster) {
      const normalized = defaultWorkspaceData();
      normalized.users = await usersToClientShapeAll();
      return res.json(normalized);
    }

    const tenantRoot = resolveTenantRoot(req);
    console.log(`GET /workspace uid=${req.user.userId} tenantRoot=${tenantRoot}`);

    const ws = await ensureWorkspaceForTenantRoot(tenantRoot);
    if (!ws) {
      console.error('GET /workspace: ensureWorkspaceForTenantRoot returned null for', tenantRoot);
      return res.status(500).json({ error: 'Failed to load workspace' });
    }

    const normalized = normalizeWorkspacePayload(ws.data);
    const users = await loadTenantUsers(tenantRoot, req.user.userId);
    console.log(`GET /workspace uid=${req.user.userId} tenantRoot=${tenantRoot} users=${users.length}`);
    normalized.users = users;

    const isAdmin = req.user.role === 'admin';
    if (!isAdmin) {
      const adminDoc = await User.findOne({
        $or: [{ userId: tenantRoot }, { tenantRootUserId: tenantRoot, role: 'admin' }],
        isMaster: { $ne: true },
      }).lean();
      const primaryAdminName = adminDoc ? (adminDoc.name || adminDoc.email) : '';
      if (Array.isArray(normalized.tasks)) {
        normalized.tasks = normalized.tasks.map((t) => ({
          ...t,
          _assignedByAdmin: primaryAdminName,
          _sourceWorkspace: tenantRoot,
        }));
      }
      const shared = await loadSharedTasks(req.user.userId);
      if (shared.length > 0) {
        normalized.tasks = [...(normalized.tasks || []), ...shared];
      }
    }

    try {
      const dataToSave = { ...normalized };
      if (Array.isArray(dataToSave.tasks)) {
        dataToSave.tasks = dataToSave.tasks.filter((t) => !t._sharedTask);
      }
      ws.data = dataToSave;
      ws.markModified('data');
      await ws.save();
    } catch (saveErr) {
      console.error('GET /workspace: could not persist into ws.data:', saveErr.message);
    }

    return res.json(normalized);
  } catch (e) {
    console.error('GET /workspace CRASH:', e);
    return res.status(500).json({ error: 'Failed to load workspace' });
  }
});

router.put('/', authMiddleware, async (req, res) => {
  try {
    if (req.user.isMaster) {
      return res.status(403).json({ error: 'Master account cannot modify workspace data' });
    }

    const doc = await User.findOne({ userId: req.user.userId });
    if (!doc || !doc.isActive) {
      return res.status(401).json({ error: 'Unauthorized' });
    }
    const isAdmin = doc.role === 'admin';
    const tenantRoot = resolveTenantRoot(req);
    console.log(`PUT /workspace uid=${req.user.userId} tenantRoot=${tenantRoot} isAdmin=${isAdmin}`);

    const ws = await ensureWorkspaceForTenantRoot(tenantRoot);
    if (!ws) {
      return res.status(500).json({ error: 'Failed to save workspace' });
    }

    const existingNormalized = normalizeWorkspacePayload(ws.data);
    const incoming = normalizeWorkspacePayload(req.body);
    if (Array.isArray(incoming.tasks)) {
      incoming.tasks = incoming.tasks.filter((t) => !t._sharedTask);
    }

    const merged = { ...incoming };

    if (!isAdmin) {
      merged.users = await loadTenantUsers(tenantRoot, req.user.userId);
      merged.locations = existingNormalized.locations;
      merged.segregationTypes = existingNormalized.segregationTypes;
      merged.holidays = existingNormalized.holidays;
    } else {
      try {
        const previousUserIds = previousWorkspaceUserIdSet(existingNormalized.users);
        const incomingUsersMerged = await mergeIncomingUsersWithDbTenantRoster(
          incoming.users,
          tenantRoot,
          previousUserIds
        );
        await syncUsersFromClientPayload(incomingUsersMerged, { isAdmin: true, tenantRootUserId: tenantRoot });
        if (Array.isArray(incomingUsersMerged) && incomingUsersMerged.length > 0) {
          await deleteUsersNotInPayload(
            incomingUsersMerged.map((u) => u.id),
            tenantRoot,
            previousUserIds,
            false
          );
        }
      } catch (syncErr) {
        console.error('PUT /workspace: user sync error (workspace will still save):', syncErr);
      }
    }

    const users = await loadTenantUsers(tenantRoot, req.user.userId);
    console.log(`PUT /workspace uid=${req.user.userId} tenantRoot=${tenantRoot} users=${users.length}`);
    merged.users = users;

    ws.data = merged;
    ws.markModified('data');
    await ws.save();

    merged._debug = {
      tenantRoot,
      reqUserId: req.user.userId,
      usersReturned: users.length,
      userIds: users.map((u) => u.id),
    };

    return res.json(merged);
  } catch (e) {
    console.error('PUT /workspace CRASH:', e);
    return res.status(500).json({ error: 'Failed to save workspace', detail: e.message });
  }
});

router.patch('/shared-task', authMiddleware, async (req, res) => {
  try {
    const { sourceWorkspace, taskId, updates } = req.body || {};
    if (!sourceWorkspace || !taskId || !updates) {
      return res.status(400).json({ error: 'sourceWorkspace, taskId, and updates required' });
    }
    const root = Number(sourceWorkspace);
    const ws = await Workspace.findOne({ tenantRootUserId: root });
    if (!ws || !ws.data || !Array.isArray(ws.data.tasks)) {
      return res.status(404).json({ error: 'Workspace not found' });
    }
    const task = ws.data.tasks.find((t) => String(t.id) === String(taskId));
    if (!task) {
      return res.status(404).json({ error: 'Task not found' });
    }
    if (Number(task.assigned_to) !== req.user.userId) {
      return res.status(403).json({ error: 'Not assigned to you' });
    }
    const allowed = ['status', 'completion_percentage'];
    for (const key of allowed) {
      if (updates[key] !== undefined) task[key] = updates[key];
    }
    ws.markModified('data');
    await ws.save();
    return res.json({ ok: true });
  } catch (e) {
    console.error('PATCH /shared-task error:', e);
    return res.status(500).json({ error: 'Update failed' });
  }
});

module.exports = router;
