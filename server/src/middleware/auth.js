const jwt = require('jsonwebtoken');
const User = require('../models/User');
const { JWT_SECRET } = require('../config');
const { isProduction } = require('../config');
const { getSiteSettings } = require('../services/registrationPolicy');

const ACTIVITY_THROTTLE_MS = 30 * 1000;

function signToken(payload) {
  return jwt.sign(payload, JWT_SECRET, { expiresIn: '7d' });
}

function verifyToken(token) {
  try {
    return jwt.verify(token, JWT_SECRET);
  } catch {
    return null;
  }
}

function authCookieOptions() {
  return {
    path: '/',
    sameSite: isProduction ? 'none' : 'lax',
    secure: isProduction,
  };
}

function clearAuthCookie(res) {
  res.clearCookie('auth_token', authCookieOptions());
}

/**
 * Validates active session: lockout, admin disable, idle timeout, and bumps lastActivityAt (throttled).
 * @returns {Promise<{ ok: boolean, sessionIdleTimeoutMinutes?: number }>}
 */
async function validateUserSession(doc, res) {
  const s = await getSiteSettings();
  const sessionIdleTimeoutMinutes =
    s.sessionIdleTimeoutMinutes != null && !Number.isNaN(Number(s.sessionIdleTimeoutMinutes))
      ? Math.max(0, Number(s.sessionIdleTimeoutMinutes))
      : 0;

  if (!doc) {
    clearAuthCookie(res);
    res.status(401).json({ error: 'Unauthorized' });
    return { ok: false, sessionIdleTimeoutMinutes };
  }
  if (doc.loginLocked) {
    clearAuthCookie(res);
    res.status(403).json({
      error:
        'Account locked after too many failed sign-in attempts. Use Forgot password to reset and unlock.',
      code: 'LOGIN_LOCKED',
    });
    return { ok: false, sessionIdleTimeoutMinutes };
  }
  if (!doc.isActive) {
    clearAuthCookie(res);
    res.status(401).json({ error: 'Account is disabled' });
    return { ok: false, sessionIdleTimeoutMinutes };
  }

  if (sessionIdleTimeoutMinutes > 0 && doc.lastActivityAt) {
    const elapsed = Date.now() - new Date(doc.lastActivityAt).getTime();
    if (elapsed > sessionIdleTimeoutMinutes * 60 * 1000) {
      clearAuthCookie(res);
      res.status(401).json({
        error: 'Session expired due to inactivity',
        code: 'SESSION_IDLE',
      });
      return { ok: false, sessionIdleTimeoutMinutes };
    }
  }

  const now = Date.now();
  const last = doc.lastActivityAt ? new Date(doc.lastActivityAt).getTime() : 0;
  if (!doc.lastActivityAt || now - last > ACTIVITY_THROTTLE_MS) {
    doc.lastActivityAt = new Date();
    await doc.save();
  }

  return { ok: true, sessionIdleTimeoutMinutes };
}

function resolveAuthDecoded(req) {
  const raw = req.cookies?.auth_token || '';
  const bearer = req.headers.authorization?.startsWith('Bearer ')
    ? req.headers.authorization.slice(7)
    : '';
  if (raw) {
    const d = verifyToken(raw);
    if (d && d.sub !== undefined && d.sub !== null) return d;
  }
  if (bearer) {
    const d = verifyToken(bearer);
    if (d && d.sub !== undefined && d.sub !== null) return d;
  }
  return null;
}

async function authMiddleware(req, res, next) {
  try {
    const decoded = resolveAuthDecoded(req);
    if (!decoded) {
      return res.status(401).json({ error: 'Unauthorized' });
    }
    const userId = typeof decoded.sub === 'number' ? decoded.sub : parseInt(String(decoded.sub), 10);
    if (Number.isNaN(userId)) {
      return res.status(401).json({ error: 'Invalid token' });
    }
    const doc = await User.findOne({ userId });
    if (!doc) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    const sess = await validateUserSession(doc, res);
    if (!sess.ok) {
      return;
    }

    const rawTr = doc.tenantRootUserId;
    let tenantRoot;
    if (rawTr != null && rawTr !== '' && !Number.isNaN(Number(rawTr))) {
      tenantRoot = Number(rawTr);
    } else {
      tenantRoot = doc.userId;
    }

    req.user = {
      userId: doc.userId,
      role: doc.role,
      isMaster: !!doc.isMaster,
      tenantRootUserId: tenantRoot,
    };
    req.sessionIdleTimeoutMinutes = sess.sessionIdleTimeoutMinutes;
    next();
  } catch (err) {
    console.error('authMiddleware error:', err);
    return res.status(500).json({ error: 'Authentication error' });
  }
}

function optionalAuth(req, res, next) {
  const decoded = resolveAuthDecoded(req);
  if (decoded && decoded.sub !== undefined && decoded.sub !== null) {
    req.user = {
      userId: typeof decoded.sub === 'number' ? decoded.sub : parseInt(String(decoded.sub), 10),
      role: decoded.role,
      isMaster: !!decoded.isMaster,
    };
  }
  next();
}

module.exports = {
  authMiddleware,
  optionalAuth,
  signToken,
  verifyToken,
  resolveAuthDecoded,
  validateUserSession,
  clearAuthCookie,
  authCookieOptions,
};
