const express = require('express');
const bcrypt = require('bcryptjs');
const User = require('../models/User');
const { signToken } = require('../middleware/auth');
const { isProduction } = require('../config');
const { usersToClientShape } = require('../services/userSync');

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
    const countNonMaster = await User.countDocuments({ isMaster: { $ne: true } });
    const role = countNonMaster === 0 ? 'admin' : 'user';
    const userId = Date.now();
    const passwordHash = await bcrypt.hash(String(password), 12);
    const doc = await User.create({
      userId,
      email: em,
      name: String(name).trim(),
      passwordHash,
      role,
      isActive: true,
      isMaster: false,
    });
    const token = signToken({
      sub: doc.userId,
      role: doc.role,
      isMaster: doc.isMaster,
    });
    res.cookie('auth_token', token, cookieOptions());
    return res.status(201).json({ user: publicUser(doc) });
  } catch (e) {
    console.error(e);
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
    return res.json({ user: publicUser(doc) });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Login failed' });
  }
});

router.post('/logout', (_req, res) => {
  res.clearCookie('auth_token', { path: '/', sameSite: isProduction ? 'none' : 'lax', secure: isProduction });
  res.json({ ok: true });
});

router.get('/me', async (req, res) => {
  const raw = req.cookies?.auth_token || '';
  const bearer = req.headers.authorization?.startsWith('Bearer ')
    ? req.headers.authorization.slice(7)
    : '';
  const token = raw || bearer;
  if (!token) {
    return res.status(401).json({ error: 'Not logged in' });
  }
  const { verifyToken } = require('../middleware/auth');
  const decoded = verifyToken(token);
  if (!decoded || decoded.sub === undefined || decoded.sub === null) {
    return res.status(401).json({ error: 'Invalid session' });
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
  const raw = req.cookies?.auth_token || '';
  const bearer = req.headers.authorization?.startsWith('Bearer ')
    ? req.headers.authorization.slice(7)
    : '';
  const token = raw || bearer;
  if (!token) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  const { verifyToken } = require('../middleware/auth');
  const decoded = verifyToken(token);
  if (!decoded || !decoded.isMaster) {
    return res.status(403).json({ error: 'Master only' });
  }
  const list = await usersToClientShape();
  return res.json(list);
});

module.exports = router;
