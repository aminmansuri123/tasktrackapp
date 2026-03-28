const express = require('express');
const bcrypt = require('bcryptjs');
const User = require('../models/User');
const { authMiddleware } = require('../middleware/auth');
const requireMaster = require('../middleware/requireMaster');
const { usersToClientShape } = require('../services/userSync');
const Workspace = require('../models/Workspace');
const { normalizeWorkspacePayload } = require('../services/defaultWorkspace');

const router = express.Router();

router.post('/users/:userId/password', authMiddleware, requireMaster, async (req, res) => {
  try {
    const userId = parseInt(req.params.userId, 10);
    const { newPassword } = req.body || {};
    if (Number.isNaN(userId) || !newPassword || String(newPassword).length < 6) {
      return res.status(400).json({ error: 'Invalid user or password (min 6 characters)' });
    }
    const target = await User.findOne({ userId });
    if (!target) {
      return res.status(404).json({ error: 'User not found' });
    }
    target.passwordHash = await bcrypt.hash(String(newPassword), 12);
    await target.save();
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Password reset failed' });
  }
});

router.post('/resync-workspace-users', authMiddleware, requireMaster, async (_req, res) => {
  try {
    const ws = await Workspace.findOne({ name: 'default' });
    if (!ws) {
      return res.json({ ok: true });
    }
    const normalized = normalizeWorkspacePayload(ws.data);
    normalized.users = await usersToClientShape();
    ws.data = normalized;
    ws.markModified('data');
    await ws.save();
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Resync failed' });
  }
});

module.exports = router;
