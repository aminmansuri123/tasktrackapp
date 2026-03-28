function requireMaster(req, res, next) {
  if (!req.user || !req.user.isMaster) {
    return res.status(403).json({ error: 'Master access required' });
  }
  next();
}

module.exports = requireMaster;
