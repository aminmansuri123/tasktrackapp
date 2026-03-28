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
  if (tenantRootUserId == null || Number.isNaN(tenantRootUserId)) {
    return [];
  }
  const list = await User.find({
    tenantRootUserId,
    isMaster: { $ne: true },
  })
    .sort({ userId: 1 })
    .lean();
  return list.map((u) => ({
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
    const userId = typeof u.id === 'number' ? u.id : parseInt(u.id, 10);
    if (!email || Number.isNaN(userId)) continue;

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
      await User.create({
        userId,
        email,
        name: u.name || email,
        passwordHash,
        role: isMaster ? 'admin' : u.role === 'admin' ? 'admin' : 'user',
        isActive: u.is_active !== false,
        isMaster,
        tenantRootUserId: isMaster ? null : tenantRootUserId,
      });
    }
  }
}

async function deleteUsersNotInPayload(allowedUserIds, tenantRootUserId) {
  if (tenantRootUserId == null) return;
  const ids = new Set(allowedUserIds.filter((id) => typeof id === 'number' && !Number.isNaN(id)));
  const all = await User.find({ tenantRootUserId });
  for (const u of all) {
    if (u.isMaster) continue;
    if (!ids.has(u.userId)) {
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
};
