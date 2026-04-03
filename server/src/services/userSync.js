const bcrypt = require('bcryptjs');
const User = require('../models/User');
const { MASTER_EMAIL } = require('../config');
const { formatLastLoginAtDisplay } = require('../lib/lastLoginFormat');

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
  const list = await User.find({
    isMaster: { $ne: true },
    $or: [
      { tenantRootUserId: root },
      { userId: root },
      { sharedWithTenants: root },
    ],
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
    enabledFeatures: Array.isArray(u.enabledFeatures) ? u.enabledFeatures : [],
    isShared: Array.isArray(u.sharedWithTenants) && u.sharedWithTenants.includes(root) && u.tenantRootUserId !== root,
    last_login_at: formatLastLoginAtDisplay(u.lastLoginAt),
  }));
}

/** All accounts (for master password tools). */
async function usersToClientShapeAll() {
  const list = await User.find().sort({ userId: 1 }).lean();
  const byId = new Map(list.map((u) => [u.userId, u]));
  return list.map((u) => {
    let adminRoot = null;
    if (u.isMaster) {
      adminRoot = null;
    } else if (u.tenantRootUserId != null && u.tenantRootUserId !== '' && !Number.isNaN(Number(u.tenantRootUserId))) {
      adminRoot = Number(u.tenantRootUserId);
    } else if (u.role === 'admin') {
      adminRoot = u.userId;
    }
    const rootUser = adminRoot != null ? byId.get(adminRoot) : null;
    let tenantAdminLabel = '—';
    if (u.isMaster) {
      tenantAdminLabel = '—';
    } else if (adminRoot != null) {
      tenantAdminLabel = rootUser
        ? `${rootUser.name || ''} (${rootUser.email || ''})`.trim() || `Account admin ID ${adminRoot}`
        : `Org root ${adminRoot}`;
    }
    return {
      id: u.userId,
      email: u.email,
      name: u.name,
      role: u.role,
      is_active: u.isActive,
      isMaster: u.isMaster,
      tenantRootUserId: u.tenantRootUserId,
      tenant_admin_root_id: adminRoot,
      tenant_admin_label: tenantAdminLabel,
      enabledFeatures: Array.isArray(u.enabledFeatures) ? u.enabledFeatures : [],
      sharedWithTenants: Array.isArray(u.sharedWithTenants) ? u.sharedWithTenants : [],
      last_login_at: formatLastLoginAtDisplay(u.lastLoginAt),
    };
  });
}

async function syncUsersFromClientPayload(usersPayload, { isAdmin, tenantRootUserId }) {
  if (!isAdmin || !Array.isArray(usersPayload) || tenantRootUserId == null) {
    return;
  }

  for (const u of usersPayload) {
    try {
      if (!u || typeof u.email !== 'string') continue;
      const email = u.email.toLowerCase().trim();
      const userId = coerceTenantUserId(u.id);
      if (!email || userId == null) continue;

      const existing = await User.findOne({ $or: [{ userId }, { email }] });
      const passwordPlain = typeof u.password === 'string' && u.password.length > 0 ? u.password : null;

      if (existing) {
        if (existing.isMaster) continue;
        const exRoot =
          existing.tenantRootUserId != null && existing.tenantRootUserId !== ''
            ? Number(existing.tenantRootUserId)
            : null;
        const tRoot = Number(tenantRootUserId);
        if (
          exRoot != null &&
          Number.isFinite(exRoot) &&
          Number.isFinite(tRoot) &&
          exRoot !== tRoot
        ) {
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
        } catch (createErr) {
          if (createErr && createErr.code === 11000) {
            const dupKey = createErr.keyPattern ? Object.keys(createErr.keyPattern)[0] : '';
            if (dupKey === 'userId') {
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
              console.error(`syncUsers: duplicate key (${dupKey}) for ${email}, skipping`, createErr.message);
            }
          } else {
            throw createErr;
          }
        }
      }
    } catch (perUserErr) {
      console.error('syncUsers: error processing user', u && u.email, perUserErr.message || perUserErr);
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
/**
 * User ids that appeared on the last persisted workspace `data.users` roster.
 * Used to avoid deleting Mongo users who were never on that snapshot (e.g. self-registered
 * account users) when an admin's browser sends a stale PUT without them.
 */
function previousWorkspaceUserIdSet(usersArray) {
  const s = new Set();
  if (!Array.isArray(usersArray)) return s;
  for (const u of usersArray) {
    const n = coerceTenantUserId(u && u.id);
    if (n != null) s.add(n);
  }
  return s;
}

/**
 * Merges client user payload with Mongo tenant roster so users never on the persisted workspace
 * snapshot (e.g. master-moved or self-registered) are not dropped from the next admin PUT/delete pass.
 */
async function mergeIncomingUsersWithDbTenantRoster(incomingUsers, tenantRootUserId, previousWorkspaceUserIds) {
  const incoming = Array.isArray(incomingUsers) ? [...incomingUsers] : [];
  const prev = previousWorkspaceUserIds instanceof Set ? previousWorkspaceUserIds : new Set();
  const dbRoster = await usersToClientShapeForTenant(tenantRootUserId);
  const incomingIds = new Set();
  for (const u of incoming) {
    const id = coerceTenantUserId(u && u.id);
    if (id != null) incomingIds.add(id);
  }
  for (const du of dbRoster) {
    if (incomingIds.has(du.id)) continue;
    if (!prev.has(du.id)) {
      incoming.push({ ...du });
      incomingIds.add(du.id);
    }
  }
  return incoming;
}

/**
 * @param {Set<number>|null|undefined} previousWorkspaceUserIds - When replaceAllTenantUsers is false,
 *   only delete Mongo users who were listed on the persisted workspace roster and are now missing
 *   from the client payload (avoids wiping self-registered account users the admin UI never saved).
 * @param {boolean} [replaceAllTenantUsers] - True for backup restore: delete anyone not in allowed.
 */
async function deleteUsersNotInPayload(
  allowedUserIdsRaw,
  tenantRootUserId,
  previousWorkspaceUserIds,
  replaceAllTenantUsers
) {
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

  const root = Number(tenantRootUserId);
  const all = await User.find({
    isMaster: { $ne: true },
    $or: [{ tenantRootUserId: root }, { userId: root }],
  });
  for (const u of all) {
    if (u.isMaster) continue;
    if (u.userId === tenantRootUserId) continue;
    const isSharedHere = Array.isArray(u.sharedWithTenants) && u.sharedWithTenants.includes(root) && u.tenantRootUserId !== root;
    if (isSharedHere) continue;
    if (!allowed.has(u.userId)) {
      if (replaceAllTenantUsers === true) {
        await User.deleteOne({ _id: u._id });
        continue;
      }
      if (!(previousWorkspaceUserIds instanceof Set) || previousWorkspaceUserIds.size === 0) {
        continue;
      }
      if (!previousWorkspaceUserIds.has(u.userId)) {
        continue;
      }
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
  coerceTenantUserId,
  previousWorkspaceUserIdSet,
  mergeIncomingUsersWithDbTenantRoster,
};
