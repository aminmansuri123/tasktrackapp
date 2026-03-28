const jwt = require('jsonwebtoken');
const { JWT_SECRET } = require('../config');

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

function authMiddleware(req, res, next) {
  const raw = req.cookies?.auth_token || '';
  const bearer = req.headers.authorization?.startsWith('Bearer ')
    ? req.headers.authorization.slice(7)
    : '';
  const token = raw || bearer;
  if (!token) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  const decoded = verifyToken(token);
  if (!decoded || decoded.sub === undefined || decoded.sub === null) {
    return res.status(401).json({ error: 'Invalid token' });
  }
  const userId = typeof decoded.sub === 'number' ? decoded.sub : parseInt(String(decoded.sub), 10);
  if (Number.isNaN(userId)) {
    return res.status(401).json({ error: 'Invalid token' });
  }
  req.user = {
    userId,
    role: decoded.role,
    isMaster: !!decoded.isMaster,
  };
  next();
}

function optionalAuth(req, res, next) {
  const raw = req.cookies?.auth_token || '';
  const bearer = req.headers.authorization?.startsWith('Bearer ')
    ? req.headers.authorization.slice(7)
    : '';
  const token = raw || bearer;
  if (token) {
    const decoded = verifyToken(token);
    if (decoded && decoded.sub) {
      req.user = {
        userId: decoded.sub,
        role: decoded.role,
        isMaster: !!decoded.isMaster,
      };
    }
  }
  next();
}

module.exports = { authMiddleware, optionalAuth, signToken, verifyToken };
