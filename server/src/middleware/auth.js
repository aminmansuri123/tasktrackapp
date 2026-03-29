const jwt = require('jsonwebtoken');
const User = require('../models/User');
const { JWT_SECRET } = require('../config');
const { resolveWorkspaceTenantRootUserId } = require('../services/tenantRoot');

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

/** Prefer a valid cookie; if cookie is missing or expired, accept Authorization Bearer (cross-site cookie fallback). */
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
  const decoded = resolveAuthDecoded(req);
  if (!decoded) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  const userId = typeof decoded.sub === 'number' ? decoded.sub : parseInt(String(decoded.sub), 10);
  if (Number.isNaN(userId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }
  const doc = await User.findOne({ userId });
  if (!doc || !doc.isActive) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  const workspaceTenantRoot = await resolveWorkspaceTenantRootUserId(doc);
  req.user = {
    userId: doc.userId,
    role: doc.role,
    isMaster: !!doc.isMaster,
    tenantRootUserId: doc.tenantRootUserId,
    workspaceTenantRoot,
  };
  next();
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

module.exports = { authMiddleware, optionalAuth, signToken, verifyToken, resolveAuthDecoded };
