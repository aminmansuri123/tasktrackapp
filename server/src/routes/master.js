const express = require('express');
const bcrypt = require('bcryptjs');
const User = require('../models/User');
const { authMiddleware } = require('../middleware/auth');
const requireMaster = require('../middleware/requireMaster');
const { usersToClientShapeForTenant } = require('../services/userSync');
const Workspace = require('../models/Workspace');
const { normalizeWorkspacePayload } = require('../services/defaultWorkspace');
const { getSiteSettings, sanitizePolicyBody } = require('../services/registrationPolicy');
const { allocateUniqueUserId } = require('../services/userSync');
const { resolveTenantRootFromAdminPicker } = require('../services/tenantRoot');
const { ensureWorkspaceForTenantRoot } = require('../services/ensureWorkspace');

const router = express.Router();

function masterUserSummary(doc) {
  return {
    id: doc.userId,
    email: doc.email,
    name: doc.name,
    role: doc.role,
    tenantRootUserId: doc.tenantRootUserId,
    is_active: doc.isActive,
    isMaster: doc.isMaster,
  };
}

router.post('/users', authMiddleware, requireMaster, async (req, res) => {
  try {
    const { name, email, password, accountType, orgAdminUserId } = req.body || {};
    if (!name || !email || !password || String(password).length < 6) {
      return res.status(400).json({ error: 'Name, email, and password (min 6 characters) required' });
    }
    const em = String(email).toLowerCase().trim();
    if (await User.findOne({ email: em })) {
      return res.status(400).json({ error: 'Email already registered' });
    }
    const userId = await allocateUniqueUserId(Date.now());
    const passwordHash = await bcrypt.hash(String(password), 12);
    const type = accountType === 'team_user' ? 'team_user' : 'org_admin';

    if (type === 'team_user') {
      const raw = orgAdminUserId;
      const pickerId = typeof raw === 'number' ? raw : parseInt(String(raw), 10);
      if (Number.isNaN(pickerId)) {
        return res.status(400).json({ error: 'Select an account admin' });
      }
      const tenantRoot = await resolveTenantRootFromAdminPicker(pickerId);
      if (tenantRoot == null) {
        return res.status(400).json({ error: 'Invalid account admin' });
      }
      const doc = await User.create({
        userId,
        email: em,
        name: String(name).trim(),
        passwordHash,
        role: 'user',
        isActive: true,
        isMaster: false,
        tenantRootUserId: tenantRoot,
      });
      return res.status(201).json({ ok: true, user: masterUserSummary(doc) });
    }

    const doc = await User.create({
      userId,
      email: em,
      name: String(name).trim(),
      passwordHash,
      role: 'admin',
      isActive: true,
      isMaster: false,
      tenantRootUserId: userId,
    });
    await ensureWorkspaceForTenantRoot(userId);
    return res.status(201).json({ ok: true, user: masterUserSummary(doc) });
  } catch (e) {
    console.error(e);
    if (e && e.code === 11000) {
      return res.status(400).json({ error: 'Email or user id already in use' });
    }
    return res.status(500).json({ error: 'Could not create user' });
  }
});

router.delete('/users/:userId', authMiddleware, requireMaster, async (req, res) => {
  try {
    const userId = parseInt(req.params.userId, 10);
    if (Number.isNaN(userId)) {
      return res.status(400).json({ error: 'Invalid user' });
    }
    const target = await User.findOne({ userId });
    if (!target) {
      return res.status(404).json({ error: 'User not found' });
    }
    if (target.isMaster) {
      return res.status(403).json({ error: 'Cannot delete master account' });
    }
    await User.deleteOne({ userId });
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Delete failed' });
  }
});

router.patch('/users/:userId', authMiddleware, requireMaster, async (req, res) => {
  try {
    const userId = parseInt(req.params.userId, 10);
    if (Number.isNaN(userId)) {
      return res.status(400).json({ error: 'Invalid user' });
    }
    const target = await User.findOne({ userId });
    if (!target) {
      return res.status(404).json({ error: 'User not found' });
    }
    if (target.isMaster) {
      return res.status(403).json({ error: 'Cannot modify master account' });
    }

    const { role, orgAdminUserId } = req.body || {};
    if (role === 'admin' || role === 'user') {
      target.role = role;
    }

    if (orgAdminUserId !== undefined && orgAdminUserId !== null && orgAdminUserId !== '') {
      const pickerId =
        typeof orgAdminUserId === 'number' ? orgAdminUserId : parseInt(String(orgAdminUserId), 10);
      if (Number.isNaN(pickerId)) {
        return res.status(400).json({ error: 'Invalid account admin' });
      }
      const tenantRoot = await resolveTenantRootFromAdminPicker(pickerId);
      if (tenantRoot == null) {
        return res.status(400).json({ error: 'Invalid account admin' });
      }
      target.tenantRootUserId = tenantRoot;
    }

    await target.save();
    return res.json({ ok: true, user: masterUserSummary(target) });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Update failed' });
  }
});

router.get('/registration-policy', authMiddleware, requireMaster, async (_req, res) => {
  try {
    const s = await getSiteSettings();
    return res.json({
      registrationMode: s.registrationMode,
      allowedEmails: s.allowedEmails || [],
      allowedDomains: s.allowedDomains || [],
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to load registration policy' });
  }
});

router.put('/registration-policy', authMiddleware, requireMaster, async (req, res) => {
  try {
    const parsed = sanitizePolicyBody(req.body || {});
    if (parsed.error) {
      return res.status(400).json({ error: parsed.error });
    }
    const s = await getSiteSettings();
    s.registrationMode = parsed.registrationMode;
    s.allowedEmails = parsed.allowedEmails;
    s.allowedDomains = parsed.allowedDomains;
    await s.save();
    return res.json({
      registrationMode: s.registrationMode,
      allowedEmails: s.allowedEmails,
      allowedDomains: s.allowedDomains,
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to save registration policy' });
  }
});

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

router.post('/users/:userId/active', authMiddleware, requireMaster, async (req, res) => {
  try {
    const userId = parseInt(req.params.userId, 10);
    const { active } = req.body || {};
    if (Number.isNaN(userId) || typeof active !== 'boolean') {
      return res.status(400).json({ error: 'Invalid user or active flag' });
    }
    const target = await User.findOne({ userId });
    if (!target) {
      return res.status(404).json({ error: 'User not found' });
    }
    if (target.isMaster) {
      return res.status(403).json({ error: 'Cannot change master account status' });
    }
    target.isActive = active;
    await target.save();
    return res.json({ ok: true, is_active: target.isActive });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Update failed' });
  }
});

router.post('/resync-workspace-users', authMiddleware, requireMaster, async (_req, res) => {
  try {
    const workspaces = await Workspace.find({ tenantRootUserId: { $ne: null } });
    // eslint-disable-next-line no-restricted-syntax
    for (const ws of workspaces) {
      const root = ws.tenantRootUserId;
      const normalized = normalizeWorkspacePayload(ws.data);
      normalized.users = await usersToClientShapeForTenant(root);
      ws.data = normalized;
      ws.markModified('data');
      // eslint-disable-next-line no-await-in-loop
      await ws.save();
    }
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Resync failed' });
  }
});

module.exports = router;
