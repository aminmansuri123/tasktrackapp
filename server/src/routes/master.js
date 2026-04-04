const express = require('express');
const bcrypt = require('bcryptjs');
const User = require('../models/User');
const { authMiddleware } = require('../middleware/auth');
const requireMaster = require('../middleware/requireMaster');
const { usersToClientShapeForTenant } = require('../services/userSync');
const Workspace = require('../models/Workspace');
const { normalizeWorkspacePayload } = require('../services/defaultWorkspace');
const {
  getSiteSettings,
  sanitizePolicyBody,
  sanitizeBlockedLists,
  sanitizeReportToOptions,
} = require('../services/registrationPolicy');
const { allocateUniqueUserId } = require('../services/userSync');
const { resolveTenantRootFromAdminPicker } = require('../services/tenantRoot');
const { ensureWorkspaceForTenantRoot } = require('../services/ensureWorkspace');
const ApprovalRequest = require('../models/ApprovalRequest');
const { normalizeEmailEntry } = require('../services/registrationPolicy');
const { isEmailEnabled, sendAccountCreatedEmail } = require('../services/emailService');

const router = express.Router();

/** Keeps workspace `data.users` aligned with the User collection after org moves. */
async function reconcileWorkspaceEmbeddedUsersForTenant(tenantRootUserId) {
  const root = Number(tenantRootUserId);
  if (!Number.isFinite(root)) return;
  await ensureWorkspaceForTenantRoot(root);
  const ws = await Workspace.findOne({ tenantRootUserId: root });
  if (!ws) return;
  const normalized = normalizeWorkspacePayload(ws.data);
  normalized.users = await usersToClientShapeForTenant(root);
  ws.data = normalized;
  ws.markModified('data');
  await ws.save();
}

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
      await ensureWorkspaceForTenantRoot(tenantRoot);
      await reconcileWorkspaceEmbeddedUsersForTenant(tenantRoot);
      if (isEmailEnabled()) {
        void sendAccountCreatedEmail(doc.email, doc.name || doc.email, 'admin_created').catch((e) =>
          console.error('Master create user welcome email:', e.message)
        );
      }
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
    await reconcileWorkspaceEmbeddedUsersForTenant(userId);
    if (isEmailEnabled()) {
      void sendAccountCreatedEmail(doc.email, doc.name || doc.email, 'admin_created').catch((e) =>
        console.error('Master create user welcome email:', e.message)
      );
    }
    return res.status(201).json({ ok: true, user: masterUserSummary(doc) });
  } catch (e) {
    console.error(e);
    if (e && e.code === 11000) {
      return res.status(400).json({ error: 'Email or user id already in use' });
    }
    return res.status(500).json({ error: 'Could not create user' });
  }
});

router.get('/users/:userId/linked-data', authMiddleware, requireMaster, async (req, res) => {
  try {
    const userId = parseInt(req.params.userId, 10);
    if (Number.isNaN(userId)) {
      return res.status(400).json({ error: 'Invalid user' });
    }
    const target = await User.findOne({ userId });
    if (!target) {
      return res.status(404).json({ error: 'User not found' });
    }

    const isOrgAdmin = target.role === 'admin' && !target.isMaster;
    const tenantRoot = target.tenantRootUserId != null ? Number(target.tenantRootUserId) : target.userId;
    let tasksAssigned = 0;
    let tasksCreated = 0;
    let orgUserCount = 0;

    if (isOrgAdmin && target.userId === tenantRoot) {
      orgUserCount = await User.countDocuments({
        isMaster: { $ne: true },
        tenantRootUserId: tenantRoot,
        userId: { $ne: target.userId },
      });
    }

    const ws = await Workspace.findOne({ tenantRootUserId: tenantRoot });
    if (ws && ws.data && Array.isArray(ws.data.tasks)) {
      for (const t of ws.data.tasks) {
        if (Number(t.assigned_to) === userId) tasksAssigned++;
        if (Number(t.created_by) === userId) tasksCreated++;
      }
    }

    return res.json({
      userId,
      name: target.name,
      email: target.email,
      role: target.role,
      isOrgAdmin,
      isOrgOwner: isOrgAdmin && target.userId === tenantRoot,
      orgUserCount,
      tasksAssigned,
      tasksCreated,
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to check linked data' });
  }
});

/** Remove a user or delegated admin from their organisation (no data merge). Cannot unlink the workspace owner. */
router.post('/users/:userId/unlink-org', authMiddleware, requireMaster, async (req, res) => {
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

    const uid = Number(target.userId);
    const wsAtSelf = await Workspace.findOne({ tenantRootUserId: uid }).select('_id').lean();
    if (wsAtSelf && target.role === 'admin' && !target.isMaster) {
      return res.status(400).json({
        error:
          'Cannot unlink the account admin that owns this workspace. Delete the user or transfer the organisation first.',
      });
    }

    const trRaw = target.tenantRootUserId;
    const hasTr = trRaw != null && trRaw !== '' && !Number.isNaN(Number(trRaw));
    const previousOrg = hasTr ? Number(trRaw) : null;

    const update = { tenantRootUserId: null };
    if (target.role === 'admin' && hasTr && Number(trRaw) !== uid) {
      update.role = 'user';
    }

    await User.updateOne({ userId: uid }, { $set: update });

    if (previousOrg != null && Number.isFinite(previousOrg)) {
      await reconcileWorkspaceEmbeddedUsersForTenant(previousOrg);
    }

    return res.json({
      ok: true,
      message: 'User removed from the organisation. They can be linked again later if needed.',
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Unlink failed' });
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

    const tenantRoot = target.tenantRootUserId != null ? Number(target.tenantRootUserId) : target.userId;
    const isOrgOwner = target.role === 'admin' && target.userId === tenantRoot;

    const ws = await Workspace.findOne({ tenantRootUserId: tenantRoot });
    if (ws && ws.data && Array.isArray(ws.data.tasks)) {
      ws.data.tasks = ws.data.tasks.filter(
        (t) => Number(t.assigned_to) !== userId && Number(t.created_by) !== userId
      );
      if (Array.isArray(ws.data.users)) {
        ws.data.users = ws.data.users.filter((u) => Number(u.id) !== userId);
      }
      ws.markModified('data');
      await ws.save();
    }

    if (isOrgOwner) {
      const linkedUsers = await User.find({
        isMaster: { $ne: true },
        tenantRootUserId: tenantRoot,
        userId: { $ne: userId },
      });
      for (const lu of linkedUsers) {
        await User.deleteOne({ _id: lu._id });
      }
      if (ws) {
        await Workspace.deleteOne({ _id: ws._id });
      }
    }

    await User.deleteOne({ userId });
    if (!isOrgOwner) {
      await reconcileWorkspaceEmbeddedUsersForTenant(tenantRoot);
    }
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
    let oldTenantRoot = null;
    if (orgAdminUserId !== undefined && orgAdminUserId !== null && orgAdminUserId !== '') {
      const tr = target.tenantRootUserId;
      if (tr != null && tr !== '' && !Number.isNaN(Number(tr))) {
        oldTenantRoot = Number(tr);
      }
    }

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

    if (orgAdminUserId !== undefined && orgAdminUserId !== null && orgAdminUserId !== '') {
      const newRootRaw = target.tenantRootUserId;
      const newRoot =
        newRootRaw != null && newRootRaw !== '' && !Number.isNaN(Number(newRootRaw))
          ? Number(newRootRaw)
          : null;
      const roots = new Set();
      if (oldTenantRoot != null && Number.isFinite(oldTenantRoot)) roots.add(oldTenantRoot);
      if (newRoot != null && Number.isFinite(newRoot)) roots.add(newRoot);
      for (const r of roots) {
        await reconcileWorkspaceEmbeddedUsersForTenant(r);
      }
    }

    return res.json({ ok: true, user: masterUserSummary(target) });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Update failed' });
  }
});

router.patch('/users/:userId/share', authMiddleware, requireMaster, async (req, res) => {
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
      return res.status(403).json({ error: 'Cannot share master account' });
    }
    const { sharedWithTenants } = req.body || {};
    if (!Array.isArray(sharedWithTenants)) {
      return res.status(400).json({ error: 'sharedWithTenants must be an array of admin userIds' });
    }
    const validRoots = [];
    for (const raw of sharedWithTenants) {
      const id = typeof raw === 'number' ? raw : parseInt(String(raw), 10);
      if (Number.isNaN(id)) continue;
      if (id === Number(target.tenantRootUserId)) continue;
      const admin = await User.findOne({ userId: id, role: 'admin', isMaster: { $ne: true }, isActive: true });
      if (admin) {
        const rootId = admin.tenantRootUserId != null ? Number(admin.tenantRootUserId) : admin.userId;
        if (!validRoots.includes(rootId)) validRoots.push(rootId);
      }
    }
    target.sharedWithTenants = validRoots;
    await target.save();

    const affectedRoots = new Set(validRoots);
    if (target.tenantRootUserId != null) affectedRoots.add(Number(target.tenantRootUserId));
    for (const r of affectedRoots) {
      await reconcileWorkspaceEmbeddedUsersForTenant(r);
    }

    return res.json({ ok: true, sharedWithTenants: target.sharedWithTenants });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Share update failed' });
  }
});

router.get('/registration-policy', authMiddleware, requireMaster, async (_req, res) => {
  try {
    const s = await getSiteSettings();
    return res.json({
      registrationMode: s.registrationMode,
      allowedEmails: s.allowedEmails || [],
      allowedDomains: s.allowedDomains || [],
      blockedEmails: s.blockedEmails || [],
      blockedDomains: s.blockedDomains || [],
      sessionIdleTimeoutMinutes:
        s.sessionIdleTimeoutMinutes != null ? Number(s.sessionIdleTimeoutMinutes) || 0 : 0,
      reportToOptions: Array.isArray(s.reportToOptions) ? s.reportToOptions : [],
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
    if (req.body && req.body.sessionIdleTimeoutMinutes !== undefined && req.body.sessionIdleTimeoutMinutes !== null) {
      const n = parseInt(String(req.body.sessionIdleTimeoutMinutes), 10);
      if (Number.isNaN(n) || n < 0 || n > 10080) {
        return res.status(400).json({ error: 'sessionIdleTimeoutMinutes must be between 0 and 10080 (minutes)' });
      }
      s.sessionIdleTimeoutMinutes = n;
    }
    if (req.body && (req.body.blockedEmails !== undefined || req.body.blockedDomains !== undefined)) {
      const bl = sanitizeBlockedLists(req.body || {});
      s.blockedEmails = bl.blockedEmails;
      s.blockedDomains = bl.blockedDomains;
    }
    if (req.body && req.body.reportToOptions !== undefined) {
      s.reportToOptions = sanitizeReportToOptions(req.body || {});
    }
    await s.save();
    return res.json({
      registrationMode: s.registrationMode,
      allowedEmails: s.allowedEmails,
      allowedDomains: s.allowedDomains,
      blockedEmails: s.blockedEmails || [],
      blockedDomains: s.blockedDomains || [],
      sessionIdleTimeoutMinutes:
        s.sessionIdleTimeoutMinutes != null ? Number(s.sessionIdleTimeoutMinutes) || 0 : 0,
      reportToOptions: Array.isArray(s.reportToOptions) ? s.reportToOptions : [],
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
    target.loginLocked = false;
    target.failedLoginAttempts = 0;
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

const VALID_FEATURES = ['locations', 'codeSnippets', 'intelligenceLayer', 'templateLibrary'];

router.patch('/users/:userId/features', authMiddleware, requireMaster, async (req, res) => {
  try {
    const userId = parseInt(req.params.userId, 10);
    if (Number.isNaN(userId)) {
      return res.status(400).json({ error: 'Invalid user' });
    }
    const target = await User.findOne({ userId });
    if (!target) {
      return res.status(404).json({ error: 'User not found' });
    }
    const { enabledFeatures } = req.body || {};
    if (!Array.isArray(enabledFeatures)) {
      return res.status(400).json({ error: 'enabledFeatures must be an array' });
    }
    target.enabledFeatures = enabledFeatures.filter((f) => VALID_FEATURES.includes(f));
    await target.save();
    return res.json({ ok: true, enabledFeatures: target.enabledFeatures });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Feature update failed' });
  }
});

router.get('/approval-requests', authMiddleware, requireMaster, async (_req, res) => {
  try {
    const list = await ApprovalRequest.find({ status: 'pending' }).sort({ createdAt: -1 }).lean();
    return res.json(list.map((r) => ({ id: r._id, email: r.email, name: r.name, requestedAt: r.createdAt })));
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to load approval requests' });
  }
});

router.post('/approval-requests/:id/approve', authMiddleware, requireMaster, async (req, res) => {
  try {
    const ar = await ApprovalRequest.findById(req.params.id);
    if (!ar || ar.status !== 'pending') {
      return res.status(404).json({ error: 'Request not found or already reviewed' });
    }
    ar.status = 'approved';
    ar.reviewedAt = new Date();
    await ar.save();

    const s = await getSiteSettings();
    const em = normalizeEmailEntry(ar.email);
    if (!s.allowedEmails.includes(em)) {
      s.allowedEmails.push(em);
      await s.save();
    }
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Approval failed' });
  }
});

router.post('/approval-requests/:id/reject', authMiddleware, requireMaster, async (req, res) => {
  try {
    const ar = await ApprovalRequest.findById(req.params.id);
    if (!ar || ar.status !== 'pending') {
      return res.status(404).json({ error: 'Request not found or already reviewed' });
    }
    ar.status = 'rejected';
    ar.reviewedAt = new Date();
    await ar.save();
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Rejection failed' });
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
