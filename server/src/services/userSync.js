const bcrypt = require('bcryptjs');
const User = require('../models/User');
const { MASTER_EMAIL } = require('../config');

function stripPasswordFromUsers(users) {
  if (!Array.isArray(users)) return [];
  return users.map((u) => {
    const { password, ...rest } = u;
    return rest;
  });
}

/** Users belonging to one org (excludes master). */
async function usersToClientShapeForTenant(tenantRootUserId) {
  const root = typeof tenantRootUserId === 'number' ? tenantRootUserId : parseInt(String(tenantRootUserId), 10);
  if (root == null || Number.isNaN(root)) {
    return [];
  }
  // Include rows tied to this tenant AND the org owner row (userId === root) even if
  // tenantRootUserId was null/legacy on that document — avoids empty User Management.
  const list = await User.find({
    isMaster: { $ne: true },
    $or: [{ tenantRootUserId: root }, { userId: root }],
  })
    .sort({ userId: 1 })
    .lean();
  const byId = new Map();
  for (const u of list) {
    if (!byId.has(u.userId)) byId.set(u.userId, u);
  }
  return [...byId.values()].map((u) => ({
    id: u.userId,
    email: u.email,
    name: u.name,
    role: u.role,
    is_active: u.isActive,
    isMaster: u.isMaster,
  }));
}

/** All accounts (for master password tools). */
async function usersToClientShapeAll() {
  const list = await User.find().sort({ userId: 1 }).lean();
  return list.map((u) => ({
    id: u.userId,
    email: u.email,
    name: u.name,
    role: u.role,
    is_active: u.isActive,
    isMaster: u.isMaster,
    tenantRootUserId: u.tenantRootUserId,
  }));
}

async function syncUsersFromClientPayload(usersPayload, { isAdmin, tenantRootUserId }) {
  if (!isAdmin || !Array.isArray(usersPayload) || tenantRootUserId == null) {
    return;
  }

  for (const u of usersPayload) {
    if (!u || typeof u.email !== 'string') continue;
    const email = u.email.toLowerCase().trim();
    const userId = coerceTenantUserId(u.id);
    if (!email || userId == null) continue;

    const existing = await User.findOne({ $or: [{ userId }, { email }] });
    const passwordPlain = typeof u.password === 'string' && u.password.length > 0 ? u.password : null;

    if (existing) {
      if (existing.isMaster) continue;
      if (existing.tenantRootUserId != null && existing.tenantRootUserId !== tenantRootUserId) {
        continue;
      }
      const emailTaken = await User.findOne({ email, userId: { $ne: existing.userId } });
      if (emailTaken) continue;

      existing.email = email;
      existing.name = u.name || existing.name;
      existing.role = u.role === 'admin' ? 'admin' : 'user';
      existing.isActive = u.is_active !== false;
      existing.tenantRootUserId = tenantRootUserId;
      if (existing.email === MASTER_EMAIL) {
        existing.isMaster = true;
        existing.role = 'admin';
      }
      if (passwordPlain) {
        existing.passwordHash = await bcrypt.hash(passwordPlain, 12);
      }
      await existing.save();
    } else {
      if (!passwordPlain) continue;
      const passwordHash = await bcrypt.hash(passwordPlain, 12);
      const isMaster = email === MASTER_EMAIL;
      const assignId = await allocateUniqueUserId(userId);
      try {
        await User.create({
          userId: assignId,
          email,
          name: u.name || email,
          passwordHash,
          role: isMaster ? 'admin' : u.role === 'admin' ? 'admin' : 'user',
          isActive: u.is_active !== false,
          isMaster,
          tenantRootUserId: isMaster ? null : tenantRootUserId,
        });
      } catch (e) {
        const dupKey = e && e.keyPattern ? Object.keys(e.keyPattern)[0] : '';
        if (e && e.code === 11000 && dupKey === 'userId') {
          const retryId = await allocateUniqueUserId(Date.now() + Math.floor(Math.random() * 1e6));
          await User.create({
            userId: retryId,
            email,
            name: u.name || email,
            passwordHash,
            role: isMaster ? 'admin' : u.role === 'admin' ? 'admin' : 'user',
            isActive: u.is_active !== false,
            isMaster,
            tenantRootUserId: isMaster ? null : tenantRootUserId,
          });
        } else {
          throw e;
        }
      }
    }
  }
}

/** Match client/Mongo user ids (numbers or numeric strings). */
function coerceTenantUserId(raw) {
  if (raw == null || raw === '') return null;
  const n = typeof raw === 'number' ? raw : parseInt(String(raw), 10);
  return Number.isFinite(n) && !Number.isNaN(n) ? n : null;
}

async function allocateUniqueUserId(preferred) {
  let id =
    typeof preferred === 'number' && Number.isFinite(preferred) && !Number.isNaN(preferred)
      ? preferred
      : Date.now();
  for (let attempt = 0; attempt < 100; attempt++) {
    // eslint-disable-next-line no-await-in-loop
    const clash = await User.findOne({ userId: id });
    if (!clash) return id;
    id = Date.now() + attempt * 97 + Math.floor(Math.random() * 9999);
  }
  throw new Error('Could not allocate unique userId');
}

/**
 * Remove tenant User docs not listed in the payload. Safe coercion: string ids from JSON
 * are accepted. If the payload lists users but none parse to valid ids, skip (avoid wiping
 * the tenant). Tenant root user is always kept.
 */
async function deleteUsersNotInPayload(allowedUserIdsRaw, tenantRootUserId) {
  if (tenantRootUserId == null || !Number.isFinite(tenantRootUserId)) return;
  if (!Array.isArray(allowedUserIdsRaw)) return;

  const allowed = new Set();
  for (const raw of allowedUserIdsRaw) {
    const n = coerceTenantUserId(raw);
    if (n != null) allowed.add(n);
  }

  if (allowedUserIdsRaw.length > 0 && allowed.size === 0) {
    console.warn('deleteUsersNotInPayload: skipped — no valid user ids in payload (would have deleted all tenant users)');
    return;
  }
  if (allowed.size === 0) return;

  allowed.add(tenantRootUserId);

  const all = await User.find({ tenantRootUserId });
  for (const u of all) {
    if (u.isMaster) continue;
    if (u.userId === tenantRootUserId) continue;
    if (!allowed.has(u.userId)) {
      await User.deleteOne({ _id: u._id });
    }
  }
}

module.exports = {
  stripPasswordFromUsers,
  usersToClientShapeForTenant,
  usersToClientShapeAll,
  syncUsersFromClientPayload,
  deleteUsersNotInPayload,
  allocateUniqueUserId,
};
