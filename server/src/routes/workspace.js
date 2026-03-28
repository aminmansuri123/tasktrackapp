const express = require('express');
const User = require('../models/User');
const Workspace = require('../models/Workspace');
const { authMiddleware } = require('../middleware/auth');
const { defaultWorkspaceData, normalizeWorkspacePayload } = require('../services/defaultWorkspace');
const {
  syncUsersFromClientPayload,
  deleteUsersNotInPayload,
  usersToClientShape,
} = require('../services/userSync');

const router = express.Router();

router.get('/', authMiddleware, async (req, res) => {
  try {
    let ws = await Workspace.findOne({ name: 'default' });
    if (!ws) {
      const data = defaultWorkspaceData();
      data.users = await usersToClientShape();
      ws = await Workspace.create({ name: 'default', data });
    }
    const normalized = normalizeWorkspacePayload(ws.data);
    normalized.users = await usersToClientShape();
    return res.json(normalized);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to load workspace' });
  }
});

router.put('/', authMiddleware, async (req, res) => {
  try {
    const doc = await User.findOne({ userId: req.user.userId });
    if (!doc || !doc.isActive) {
      return res.status(401).json({ error: 'Unauthorized' });
    }
    const isAdmin = doc.role === 'admin';

    let ws = await Workspace.findOne({ name: 'default' });
    if (!ws) {
      const data = defaultWorkspaceData();
      ws = await Workspace.create({ name: 'default', data });
    }

    const existingNormalized = normalizeWorkspacePayload(ws.data);
    const incoming = normalizeWorkspacePayload(req.body);

    const merged = { ...incoming };

    if (!isAdmin) {
      merged.users = await usersToClientShape();
      merged.locations = existingNormalized.locations;
      merged.segregationTypes = existingNormalized.segregationTypes;
      merged.holidays = existingNormalized.holidays;
    } else {
      const incomingUsers = incoming.users;
      await syncUsersFromClientPayload(incomingUsers, { isAdmin: true });
      if (Array.isArray(incomingUsers) && incomingUsers.length > 0) {
        await deleteUsersNotInPayload(incomingUsers.map((u) => u.id));
      }
    }

    merged.users = await usersToClientShape();

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
