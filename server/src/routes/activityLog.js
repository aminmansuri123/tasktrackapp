const express = require('express');
const rateLimit = require('express-rate-limit');
const ActivityLog = require('../models/ActivityLog');
const { authMiddleware } = require('../middleware/auth');
const requireMaster = require('../middleware/requireMaster');

const router = express.Router();

const postLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 120,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many activity log writes' },
});

function tenantRootForActor(req) {
  if (req.user.isMaster) return null;
  const tr = req.user.tenantRootUserId;
  if (tr != null && !Number.isNaN(Number(tr))) return Number(tr);
  if (req.user.role === 'admin') return Number(req.user.userId);
  return null;
}

router.get('/', authMiddleware, async (req, res) => {
  try {
    if (req.user.isMaster) {
      return res.json({ items: [], message: 'Master account has no activity log entries here.' });
    }
    const limit = Math.min(200, Math.max(1, parseInt(String(req.query.limit || '80'), 10) || 80));
    const items = await ActivityLog.find({ actorUserId: req.user.userId })
      .sort({ createdAt: -1 })
      .limit(limit)
      .lean();
    return res.json({
      items: items.map((r) => ({
        id: String(r._id),
        action: r.action,
        entityType: r.entityType,
        entityId: r.entityId,
        summary: r.summary,
        createdAt: r.createdAt,
      })),
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to load activity log' });
  }
});

router.post('/', authMiddleware, postLimiter, async (req, res) => {
  try {
    if (req.user.isMaster) {
      return res.status(403).json({ error: 'Master account does not log here' });
    }
    const { action, entityType, entityId, summary } = req.body || {};
    const act = String(action || '').trim().slice(0, 64);
    if (!act) return res.status(400).json({ error: 'action required' });
    const sum = String(summary || '').trim().slice(0, 500);
    const doc = await ActivityLog.create({
      actorUserId: req.user.userId,
      tenantRootUserId: tenantRootForActor(req),
      action: act,
      entityType: String(entityType || '').trim().slice(0, 32),
      entityId: String(entityId != null ? entityId : '').trim().slice(0, 64),
      summary: sum,
    });
    return res.status(201).json({ ok: true, id: String(doc._id) });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Could not save activity' });
  }
});

router.delete('/all', authMiddleware, requireMaster, async (_req, res) => {
  try {
    const r = await ActivityLog.deleteMany({});
    return res.json({ ok: true, deletedCount: r.deletedCount });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Could not clear logs' });
  }
});

module.exports = router;
