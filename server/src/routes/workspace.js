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
} = require('../services/userSync');

const router = express.Router();

const EXPORT_VERSION = '12.0.6';

function resolveTenantRoot(req) {
  if (req.user.isMaster) return null;
  if (req.user.tenantRootUserId != null && !Number.isNaN(req.user.tenantRootUserId)) {
    return req.user.tenantRootUserId;
  }
  return req.user.userId;
}

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
    normalized.users = await usersToClientShapeForTenant(tenantRoot);
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
    await syncUsersFromClientPayload(complete.users, { isAdmin: true, tenantRootUserId: tenantRoot });
    if (Array.isArray(complete.users) && complete.users.length > 0) {
      await deleteUsersNotInPayload(complete.users.map((u) => u.id), tenantRoot);
    }
    const normalized = normalizeWorkspacePayload(ws.data);
    normalized.users = await usersToClientShapeForTenant(tenantRoot);
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
    const ws = await ensureWorkspaceForTenantRoot(tenantRoot);
    if (!ws) {
      return res.status(500).json({ error: 'Failed to load workspace' });
    }

    const normalized = normalizeWorkspacePayload(ws.data);
    normalized.users = await usersToClientShapeForTenant(tenantRoot);
    return res.json(normalized);
  } catch (e) {
    console.error(e);
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

    const ws = await ensureWorkspaceForTenantRoot(tenantRoot);
    if (!ws) {
      return res.status(500).json({ error: 'Failed to save workspace' });
    }

    const existingNormalized = normalizeWorkspacePayload(ws.data);
    const incoming = normalizeWorkspacePayload(req.body);

    const merged = { ...incoming };

    if (!isAdmin) {
      merged.users = await usersToClientShapeForTenant(tenantRoot);
      merged.locations = existingNormalized.locations;
      merged.segregationTypes = existingNormalized.segregationTypes;
      merged.holidays = existingNormalized.holidays;
    } else {
      const incomingUsers = incoming.users;
      await syncUsersFromClientPayload(incomingUsers, { isAdmin: true, tenantRootUserId: tenantRoot });
      if (Array.isArray(incomingUsers) && incomingUsers.length > 0) {
        await deleteUsersNotInPayload(
          incomingUsers.map((u) => u.id),
          tenantRoot
        );
      }
    }

    merged.users = await usersToClientShapeForTenant(tenantRoot);

    ws.data = merged;
    ws.markModified('data');
    await ws.save();

    return res.json(merged);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to save workspace' });
  }
});

module.exports = router;
