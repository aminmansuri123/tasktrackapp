const express = require('express');
const bcrypt = require('bcryptjs');
const User = require('../models/User');
const { signToken, authMiddleware, resolveAuthDecoded } = require('../middleware/auth');
const { ensureWorkspaceForTenantRoot } = require('../services/ensureWorkspace');
const { isProduction } = require('../config');
const { usersToClientShapeAll } = require('../services/userSync');
const { assertRegistrationAllowed, getSiteSettings } = require('../services/registrationPolicy');

const router = express.Router();

function cookieOptions() {
  return {
    httpOnly: true,
    secure: isProduction,
    sameSite: isProduction ? 'none' : 'lax',
    maxAge: 7 * 24 * 60 * 60 * 1000,
    path: '/',
  };
}

function publicUser(u) {
  return {
    id: u.userId,
    email: u.email,
    name: u.name,
    role: u.role,
    is_active: u.isActive,
    isMaster: u.isMaster,
  };
}

router.get('/registration-policy', async (_req, res) => {
  try {
    const s = await getSiteSettings();
    return res.json({ registrationMode: s.registrationMode || 'open' });
  } catch (e) {
    console.error(e);
    return res.json({ registrationMode: 'open' });
  }
});

router.post('/register', async (req, res) => {
  try {
    const { name, email, password } = req.body || {};
    if (!name || !email || !password || String(password).length < 6) {
      return res.status(400).json({ error: 'Invalid registration data' });
    }
    const em = String(email).toLowerCase().trim();
    const exists = await User.findOne({ email: em });
    if (exists) {
      return res.status(400).json({ error: 'Email already registered' });
    }
    const denied = await assertRegistrationAllowed(em);
    if (denied) {
      return res.status(403).json({ error: denied });
    }
    const userId = Date.now();
    const passwordHash = await bcrypt.hash(String(password), 12);
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
    const token = signToken({
      sub: doc.userId,
      role: doc.role,
      isMaster: doc.isMaster,
    });
    res.cookie('auth_token', token, cookieOptions());
    return res.status(201).json({ user: publicUser(doc), token });
  } catch (e) {
    console.error(e);
    if (e && e.code === 11000) {
      const dupKey = e.keyPattern && Object.keys(e.keyPattern)[0];
      if (dupKey === 'email') {
        return res.status(400).json({ error: 'Email already registered' });
      }
      return res.status(409).json({ error: 'Account setup conflict — try signing in' });
    }
    return res.status(500).json({ error: 'Registration failed' });
  }
});

router.post('/login', async (req, res) => {
  try {
    const { email, password } = req.body || {};
    if (!email || !password) {
      return res.status(400).json({ error: 'Email and password required' });
    }
    const em = String(email).toLowerCase().trim();
    const doc = await User.findOne({ email: em });
    if (!doc || !(await bcrypt.compare(String(password), doc.passwordHash))) {
      return res.status(401).json({ error: 'Invalid email or password' });
    }
    if (!doc.isActive) {
      return res.status(403).json({ error: 'Account is disabled' });
    }
    const token = signToken({
      sub: doc.userId,
      role: doc.role,
      isMaster: doc.isMaster,
    });
    res.cookie('auth_token', token, cookieOptions());
    return res.json({ user: publicUser(doc), token });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Login failed' });
  }
});

router.post('/logout', (_req, res) => {
  res.clearCookie('auth_token', { path: '/', sameSite: isProduction ? 'none' : 'lax', secure: isProduction });
  res.json({ ok: true });
});

router.post('/change-password', authMiddleware, async (req, res) => {
  try {
    const { currentPassword, newPassword } = req.body || {};
    if (!currentPassword || !newPassword || String(newPassword).length < 6) {
      return res.status(400).json({ error: 'Current password and new password (min 6 characters) required' });
    }
    const doc = await User.findOne({ userId: req.user.userId });
    if (!doc || !(await bcrypt.compare(String(currentPassword), doc.passwordHash))) {
      return res.status(401).json({ error: 'Current password is incorrect' });
    }
    doc.passwordHash = await bcrypt.hash(String(newPassword), 12);
    await doc.save();
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Could not update password' });
  }
});

router.get('/me', async (req, res) => {
  const decoded = resolveAuthDecoded(req);
  if (!decoded) {
    return res.status(401).json({ error: 'Not logged in' });
  }
  const userId = typeof decoded.sub === 'number' ? decoded.sub : parseInt(String(decoded.sub), 10);
  if (Number.isNaN(userId)) {
    return res.status(401).json({ error: 'Invalid session' });
  }
  const doc = await User.findOne({ userId });
  if (!doc || !doc.isActive) {
    return res.status(401).json({ error: 'Invalid session' });
  }
  return res.json(publicUser(doc));
});

router.get('/users-for-master', async (req, res) => {
  const decoded = resolveAuthDecoded(req);
  if (!decoded) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  if (!decoded.isMaster) {
    return res.status(403).json({ error: 'Master only' });
  }
  const list = await usersToClientShapeAll();
  return res.json(list);
});

module.exports = router;
