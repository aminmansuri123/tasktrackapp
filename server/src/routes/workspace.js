const express = require('express');
const rateLimit = require('express-rate-limit');
const User = require('../models/User');
const Workspace = require('../models/Workspace');
const ReminderPreference = require('../models/ReminderPreference');
const { authMiddleware } = require('../middleware/auth');
const { defaultWorkspaceData, normalizeWorkspacePayload } = require('../services/defaultWorkspace');
const { ensureWorkspaceForTenantRoot } = require('../services/ensureWorkspace');
const {
  isEmailEnabled,
  sendTestEmail,
  sendTaskAssignmentEmail,
  sendTaskRejectedEmail,
  sendAccountCreatedEmail,
  sendTaskViewSummaryEmail,
} = require('../services/emailService');
const { EMAIL_CONFIGURED } = require('../config');
const {
  syncUsersFromClientPayload,
  deleteUsersNotInPayload,
  usersToClientShapeForTenant,
  usersToClientShapeAll,
  previousWorkspaceUserIdSet,
  mergeIncomingUsersWithDbTenantRoster,
} = require('../services/userSync');
const { validateBody } = require('../middleware/validateBody');
const {
  workspacePutSchema,
  parseWorkspaceRestoreBody,
  emailTaskViewSummarySchema,
} = require('../validation/schemas');
const { getSiteSettings, sanitizeReportToOptions } = require('../services/registrationPolicy');
const { formatLastLoginAtDisplay } = require('../lib/lastLoginFormat');

const router = express.Router();

const EXPORT_VERSION = '17.2.4';

function isLegacyFlatJournal(j) {
  if (!j || typeof j !== 'object' || Array.isArray(j)) return false;
  const keys = Object.keys(j);
  if (keys.length === 0) return false;
  return keys.every((k) => /^\d{4}-\d{2}-\d{2}$/.test(k) && typeof j[k] === 'string');
}

/** Legacy flat journal { 'YYYY-MM-DD': html } → { [userId]: { 'YYYY-MM-DD': html } } */
function ensureJournalNested(journal, legacyOwnerUserId) {
  if (!journal || typeof journal !== 'object') return {};
  if (isLegacyFlatJournal(journal)) {
    return { [String(legacyOwnerUserId)]: { ...journal } };
  }
  const keys = Object.keys(journal);
  if (keys.length === 0) return {};
  const first = journal[keys[0]];
  if (first && typeof first === 'object' && !Array.isArray(first)) {
    const out = {};
    for (const k of keys) {
      const v = journal[k];
      if (v && typeof v === 'object' && !Array.isArray(v)) out[k] = { ...v };
    }
    return out;
  }
  return { [String(legacyOwnerUserId)]: {} };
}

function journalFlatForUser(nestedJournal, userId) {
  const slice = nestedJournal && nestedJournal[String(userId)];
  if (slice && typeof slice === 'object' && !Array.isArray(slice)) return { ...slice };
  return {};
}

function mergeJournalUserSlice(existingJournal, incomingFlat, userId, tenantRoot) {
  const nested = ensureJournalNested(existingJournal, tenantRoot);
  const u = String(userId);
  const next = { ...nested };
  next[u] = incomingFlat && typeof incomingFlat === 'object' ? { ...incomingFlat } : {};
  return next;
}

function filterByCreatedBy(arr, userId) {
  const uid = Number(userId);
  return (Array.isArray(arr) ? arr : []).filter(
    (item) => item && item.created_by != null && Number(item.created_by) === uid
  );
}

function mergeArrayByCreatedBy(existing, incoming, userId) {
  const uid = Number(userId);
  const ex = Array.isArray(existing) ? existing : [];
  const inc = Array.isArray(incoming) ? incoming : [];
  const rest = ex.filter((item) => {
    const cb = item && item.created_by != null ? Number(item.created_by) : null;
    return cb !== uid;
  });
  const mine = inc.filter((item) => item && Number(item.created_by) === uid);
  return [...rest, ...mine];
}

/** Migrate legacy shared journal + items without created_by (assign to tenant root). */
function migrateWorkspaceDataInPlace(ws, tenantRoot) {
  if (!ws.data || typeof ws.data !== 'object') return false;
  let changed = false;
  const d = ws.data;
  if (isLegacyFlatJournal(d.journal)) {
    d.journal = { [String(tenantRoot)]: { ...d.journal } };
    changed = true;
  }
  const listKeys = ['milestones', 'notes', 'learningNotes', 'dailyPlanner', 'templateBlocks'];
  for (const key of listKeys) {
    if (!Array.isArray(d[key])) continue;
    for (const item of d[key]) {
      if (item && item.created_by == null) {
        item.created_by = tenantRoot;
        changed = true;
      }
    }
  }
  if (Array.isArray(d.codeSnippets)) {
    for (const item of d.codeSnippets) {
      if (item && item.created_by == null) {
        item.created_by = tenantRoot;
        changed = true;
      }
    }
  }
  if (changed) ws.markModified('data');
  return changed;
}

/** Non-admin: only this user's rows + their diary slice. */
function scopeWorkspaceForTenantUser(normalized, userId, tenantRoot) {
  const nested = ensureJournalNested(normalized.journal, tenantRoot);
  return {
    ...scopePersonalListsForUser(normalized, userId),
    journal: journalFlatForUser(nested, userId),
  };
}

/** Admin GET/PUT response: personal lists scoped to the logged-in user (same as non-admin for these fields). */
function scopePersonalListsForUser(obj, userId) {
  const uid = Number(userId);
  return {
    ...obj,
    milestones: filterByCreatedBy(obj.milestones, uid),
    notes: filterByCreatedBy(obj.notes, uid),
    learningNotes: filterByCreatedBy(obj.learningNotes, uid),
    dailyPlanner: filterByCreatedBy(obj.dailyPlanner, uid),
    codeSnippets: filterByCreatedBy(obj.codeSnippets, uid),
    templateBlocks: filterByCreatedBy(obj.templateBlocks, uid),
  };
}

/**
 * Tenant admin promoted from user (userId !== org tenant root): only their assigned tasks, not the org owner's.
 */
function isDelegatedTenantAdmin(req, tenantRoot) {
  return (
    req.user.role === 'admin' &&
    !req.user.isMaster &&
    Number(req.user.userId) !== Number(tenantRoot)
  );
}

/** Merge PUT /workspace tasks so delegated admins cannot drop other users' tasks. */
function mergeTasksForDelegatedAdmin(existingTasks, incomingTasks, userId) {
  const uid = Number(userId);
  const existing = Array.isArray(existingTasks) ? existingTasks : [];
  const incoming = Array.isArray(incomingTasks) ? incomingTasks : [];
  const incomingById = new Map(incoming.map((t) => [String(t.id), t]));
  const merged = existing.map((t) => {
    const inc = incomingById.get(String(t.id));
    if (inc && Number(t.assigned_to) === uid) {
      return { ...t, ...inc, assigned_to: t.assigned_to };
    }
    return t;
  });
  const existingIds = new Set(existing.map((t) => String(t.id)));
  for (const inc of incoming) {
    if (!existingIds.has(String(inc.id)) && Number(inc.assigned_to) === uid) {
      merged.push(inc);
    }
  }
  return merged;
}

const emailTaskViewLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 20,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many email summary requests. Try again later.' },
});

/**
 * Determine workspace tenant root for the current request.
 * For org-owner admins: tenantRootUserId === userId → use it.
 * For delegated admins: tenantRootUserId !== userId → use tenantRootUserId.
 * For regular users: use tenantRootUserId.
 * Fallback: use userId.
 */
function resolveTenantRoot(req) {
  if (req.user.isMaster) return null;
  const tr = req.user.tenantRootUserId;
  if (tr != null && !Number.isNaN(Number(tr))) {
    return Number(tr);
  }
  return null;
}

/** Whether user doc belongs to this tenant workspace (member, org owner, or shared-in). */
function userInTenantWorkspace(tenantRoot, u) {
  if (!u || u.isMaster) return false;
  const root = Number(tenantRoot);
  if (!Number.isFinite(root)) return false;
  const tr = u.tenantRootUserId != null ? Number(u.tenantRootUserId) : null;
  if (tr === root) return true;
  if (Number(u.userId) === root) return true;
  if (Array.isArray(u.sharedWithTenants) && u.sharedWithTenants.map(Number).includes(root)) return true;
  return false;
}

async function loadTenantUsers(tenantRoot, currentUserId) {
  let users = [];
  try {
    users = await usersToClientShapeForTenant(tenantRoot);
  } catch (e) {
    console.error('loadTenantUsers: usersToClientShapeForTenant failed:', e.message);
  }

  const uid = Number(currentUserId);
  if (!Number.isNaN(uid) && !users.some((u) => Number(u.id) === uid)) {
    try {
      const doc = await User.findOne({ userId: uid }).lean();
      if (doc && !doc.isMaster) {
        users.push({
          id: doc.userId,
          email: doc.email,
          name: doc.name,
          role: doc.role,
          is_active: doc.isActive,
          isMaster: doc.isMaster,
          enabledFeatures: Array.isArray(doc.enabledFeatures) ? doc.enabledFeatures : [],
          last_login_at: doc.lastLoginAt ? formatLastLoginAtDisplay(doc.lastLoginAt) : '',
        });
      }
    } catch (e) {
      console.error('loadTenantUsers: fallback user lookup failed:', e.message);
    }
  }
  return users;
}

async function loadSharedTasks(userId) {
  try {
    const userDoc = await User.findOne({ userId }).lean();
    if (!userDoc || !Array.isArray(userDoc.sharedWithTenants) || userDoc.sharedWithTenants.length === 0) {
      return [];
    }
    const sharedTasks = [];
    for (const tenantId of userDoc.sharedWithTenants) {
      const ws = await Workspace.findOne({ tenantRootUserId: tenantId }).lean();
      if (!ws || !ws.data || !Array.isArray(ws.data.tasks)) continue;
      const admin = await User.findOne({
        $or: [{ userId: tenantId }, { tenantRootUserId: tenantId, role: 'admin' }],
        isMaster: { $ne: true },
      }).lean();
      const adminName = admin ? (admin.name || admin.email) : `Org ${tenantId}`;
      const tasksForUser = ws.data.tasks.filter(
        (t) => Number(t.assigned_to) === userId
      );
      for (const t of tasksForUser) {
        sharedTasks.push({
          ...t,
          _sharedTask: true,
          _sourceWorkspace: tenantId,
          _assignedByAdmin: adminName,
        });
      }
    }
    return sharedTasks;
  } catch (e) {
    console.error('loadSharedTasks error:', e.message);
    return [];
  }
}

router.get('/debug-tenant', authMiddleware, async (req, res) => {
  try {
    const tenantRoot = resolveTenantRoot(req);
    const uid = req.user.userId;

    const adminDoc = await User.findOne({ userId: uid }).lean();
    const wsDoc = await Workspace.findOne({ tenantRootUserId: tenantRoot }).lean();
    const allTenantUsers = await User.find({
      isMaster: { $ne: true },
      $or: [{ tenantRootUserId: tenantRoot }, { userId: tenantRoot }],
    }).lean();

    const allWorkspaces = await Workspace.find().select('tenantRootUserId').lean();

    return res.json({
      resolvedTenantRoot: tenantRoot,
      reqUser: req.user,
      adminDoc: adminDoc
        ? { userId: adminDoc.userId, email: adminDoc.email, role: adminDoc.role, tenantRootUserId: adminDoc.tenantRootUserId, isMaster: adminDoc.isMaster, isActive: adminDoc.isActive }
        : null,
      workspaceExists: !!wsDoc,
      workspaceTenantRootUserId: wsDoc ? wsDoc.tenantRootUserId : null,
      workspaceDataUsersCount: wsDoc && wsDoc.data ? (wsDoc.data.users || []).length : 0,
      tenantUsersFromDb: allTenantUsers.map((u) => ({ userId: u.userId, email: u.email, tenantRootUserId: u.tenantRootUserId, role: u.role })),
      allWorkspaceTenantRoots: allWorkspaces.map((w) => w.tenantRootUserId),
    });
  } catch (e) {
    console.error('debug-tenant error:', e);
    return res.status(500).json({ error: e.message });
  }
});

router.get('/backup', authMiddleware, async (req, res) => {
  try {
    if (req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin only' });
    }
    if (req.user.isMaster) {
      return res.status(403).json({ error: 'Use a tenant admin account for workspace backup' });
    }
    const tenantRoot = resolveTenantRoot(req);
    const ws = await ensureWorkspaceForTenantRoot(tenantRoot);
    if (!ws) {
      return res.status(404).json({ error: 'No workspace' });
    }
    const normalized = normalizeWorkspacePayload(ws.data);
    normalized.users = await loadTenantUsers(tenantRoot, req.user.userId);
    return res.json({
      version: EXPORT_VERSION,
      exportDate: new Date().toISOString(),
      applicationName: 'Task Management System',
      data: normalized,
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Backup failed' });
  }
});

router.post('/restore', authMiddleware, async (req, res) => {
  try {
    if (req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin only' });
    }
    if (req.user.isMaster) {
      return res.status(403).json({ error: 'Use a tenant admin account for restore' });
    }
    const tenantRoot = resolveTenantRoot(req);
    const body = req.body || {};
    const restored = parseWorkspaceRestoreBody(body);
    if (!restored.ok) {
      return res.status(400).json({ error: 'Invalid backup payload', details: restored.details });
    }
    const complete = normalizeWorkspacePayload(restored.data);
    let ws = await Workspace.findOne({ tenantRootUserId: tenantRoot });
    if (!ws) {
      try {
        ws = await Workspace.create({ tenantRootUserId: tenantRoot, data: complete });
      } catch (e) {
        if (e && e.code === 11000) {
          ws = await Workspace.findOne({ tenantRootUserId: tenantRoot });
        } else throw e;
      }
    }
    if (!ws) {
      return res.status(500).json({ error: 'Restore failed' });
    }
    ws.data = complete;
    ws.markModified('data');
    await ws.save();
    try {
      await syncUsersFromClientPayload(complete.users, { isAdmin: true, tenantRootUserId: tenantRoot });
      if (Array.isArray(complete.users) && complete.users.length > 0) {
        await deleteUsersNotInPayload(complete.users.map((u) => u.id), tenantRoot, null, true);
      }
    } catch (syncErr) {
      console.error('Restore: user sync error:', syncErr.message);
    }
    const normalized = normalizeWorkspacePayload(ws.data);
    normalized.users = await loadTenantUsers(tenantRoot, req.user.userId);
    return res.json(normalized);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Restore failed' });
  }
});

router.get('/updated-at', authMiddleware, async (req, res) => {
  try {
    if (req.user.isMaster) {
      return res.json({ updatedAt: null });
    }
    const tenantRoot = resolveTenantRoot(req);
    if (tenantRoot == null) {
      return res.json({ updatedAt: null });
    }
    const ws = await Workspace.findOne({ tenantRootUserId: tenantRoot }).select('updatedAt').lean();
    if (!ws || !ws.updatedAt) {
      return res.json({ updatedAt: null });
    }
    return res.json({ updatedAt: ws.updatedAt.toISOString() });
  } catch (e) {
    console.error('GET /workspace/updated-at', e);
    return res.status(500).json({ error: 'Failed' });
  }
});

/** Tenant admin: attach a registered account user (no org yet) by email — task assignment only; no data merge. */
router.post('/link-user-by-email', authMiddleware, async (req, res) => {
  try {
    if (req.user.isMaster) {
      return res.status(403).json({ error: 'Use a tenant admin account' });
    }
    if (req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin only' });
    }
    const tenantRoot = resolveTenantRoot(req);
    if (tenantRoot == null) {
      return res.status(403).json({ error: 'No organisation context' });
    }
    const email = String(req.body?.email || '')
      .toLowerCase()
      .trim();
    if (!email) {
      return res.status(400).json({ error: 'Email required' });
    }
    const target = await User.findOne({ email }).lean();
    if (!target) {
      return res.status(404).json({ error: 'No account with that email' });
    }
    if (target.isMaster) {
      return res.status(400).json({ error: 'Invalid target' });
    }
    if (Number(target.userId) === Number(req.user.userId)) {
      return res.status(400).json({ error: 'Cannot link yourself' });
    }
    const tr = target.tenantRootUserId;
    const hasRoot = tr != null && tr !== '' && !Number.isNaN(Number(tr));
    if (hasRoot) {
      if (Number(tr) === Number(tenantRoot)) {
        return res.json({ ok: true, message: 'User is already in this organisation' });
      }
      return res.status(400).json({ error: 'That user already belongs to another organisation' });
    }
    await User.updateOne({ userId: target.userId }, { $set: { tenantRootUserId: tenantRoot } });
    return res.json({ ok: true, message: 'User linked. They can refresh or sign in again to load your workspace.' });
  } catch (e) {
    console.error('POST /workspace/link-user-by-email', e);
    return res.status(500).json({ error: 'Failed to link user' });
  }
});

router.get('/', authMiddleware, async (req, res) => {
  try {
    if (req.user.isMaster) {
      const normalized = defaultWorkspaceData();
      normalized.users = await usersToClientShapeAll();
      return res.json(normalized);
    }

    const tenantRoot = resolveTenantRoot(req);
    if (tenantRoot == null) {
      if (req.user.role === 'user') {
        const udoc = await User.findOne({ userId: req.user.userId }).lean();
        if (!udoc) {
          return res.status(401).json({ error: 'User not found' });
        }
        const shell = defaultWorkspaceData();
        shell.tasks = [];
        shell.locations = [];
        shell.holidays = [];
        shell.milestones = [];
        shell.notes = [];
        shell.learningNotes = [];
        shell.dailyPlanner = [];
        shell.codeSnippets = [];
        shell.templateBlocks = [];
        shell.users = [
          {
            id: udoc.userId,
            email: udoc.email,
            name: udoc.name,
            role: udoc.role,
            is_active: udoc.isActive,
            isMaster: false,
            enabledFeatures: Array.isArray(udoc.enabledFeatures) ? udoc.enabledFeatures : [],
            last_login_at: formatLastLoginAtDisplay(udoc.lastLoginAt),
          },
        ];
        shell._pendingTenantLink = true;
        return res.json(shell);
      }
      return res.status(403).json({ error: 'No organisation workspace for this account', code: 'NO_TENANT' });
    }

    console.log(`GET /workspace uid=${req.user.userId} tenantRoot=${tenantRoot}`);

    const ws = await ensureWorkspaceForTenantRoot(tenantRoot);
    if (!ws) {
      console.error('GET /workspace: ensureWorkspaceForTenantRoot returned null for', tenantRoot);
      return res.status(500).json({ error: 'Failed to load workspace' });
    }

    try {
      if (migrateWorkspaceDataInPlace(ws, tenantRoot)) {
        await ws.save();
      }
    } catch (migErr) {
      console.error('GET /workspace migrate:', migErr.message);
    }

    const normalized = normalizeWorkspacePayload(ws.data);
    const users = await loadTenantUsers(tenantRoot, req.user.userId);
    console.log(`GET /workspace uid=${req.user.userId} tenantRoot=${tenantRoot} users=${users.length}`);
    normalized.users = users;

    const isAdmin = req.user.role === 'admin';
    if (!isAdmin) {
      const adminDoc = await User.findOne({
        $or: [{ userId: tenantRoot }, { tenantRootUserId: tenantRoot, role: 'admin' }],
        isMaster: { $ne: true },
      }).lean();
      const primaryAdminName = adminDoc ? (adminDoc.name || adminDoc.email) : '';
      if (Array.isArray(normalized.tasks)) {
        normalized.tasks = normalized.tasks.map((t) => ({
          ...t,
          _assignedByAdmin: primaryAdminName,
          _sourceWorkspace: tenantRoot,
        }));
      }
      const shared = await loadSharedTasks(req.user.userId);
      if (shared.length > 0) {
        normalized.tasks = [...(normalized.tasks || []), ...shared];
      }
    }

    if (isDelegatedTenantAdmin(req, tenantRoot) && Array.isArray(normalized.tasks)) {
      normalized.tasks = normalized.tasks.filter((t) => Number(t.assigned_to) === req.user.userId);
    }

    // Never persist GET response to ws.data. A previous version saved here after filtering tasks for
    // delegated admins, which overwrote the real workspace in Mongo with a partial task list.

    const nestedJ = ensureJournalNested(normalized.journal, tenantRoot);
    let out = { ...normalized, journal: nestedJ };
    if (req.user.role !== 'admin') {
      out = scopeWorkspaceForTenantUser(out, req.user.userId, tenantRoot);
    } else {
      out = scopePersonalListsForUser(
        { ...normalized, journal: journalFlatForUser(nestedJ, req.user.userId) },
        req.user.userId
      );
    }
    if (ws.updatedAt) {
      out._workspaceUpdatedAt = ws.updatedAt.toISOString();
    }
    try {
      const site = await getSiteSettings();
      const siteRt = Array.isArray(site.reportToOptions) ? site.reportToOptions : [];
      const tenantSan = sanitizeReportToOptions({ reportToOptions: normalized.reportToOptions });
      if (isAdmin) {
        out.tenantReportToOptions = tenantSan;
      }
      out.reportToOptions = tenantSan.length > 0 ? tenantSan : siteRt;
    } catch (_) {
      out.reportToOptions = [];
    }
    return res.json(out);
  } catch (e) {
    console.error('GET /workspace CRASH:', e);
    return res.status(500).json({ error: 'Failed to load workspace' });
  }
});

router.put('/', authMiddleware, validateBody(workspacePutSchema), async (req, res) => {
  try {
    if (req.user.isMaster) {
      return res.status(403).json({ error: 'Master account cannot modify workspace data' });
    }

    if (resolveTenantRoot(req) == null) {
      return res.status(403).json({
        error: 'Account not linked to an organisation yet. Ask your administrator to add your email under Settings.',
        code: 'NO_TENANT',
      });
    }

    const doc = await User.findOne({ userId: req.user.userId });
    if (!doc || !doc.isActive) {
      return res.status(401).json({ error: 'Unauthorized' });
    }
    const isAdmin = doc.role === 'admin';
    const tenantRoot = resolveTenantRoot(req);
    console.log(`PUT /workspace uid=${req.user.userId} tenantRoot=${tenantRoot} isAdmin=${isAdmin}`);

    const ws = await ensureWorkspaceForTenantRoot(tenantRoot);
    if (!ws) {
      return res.status(500).json({ error: 'Failed to save workspace' });
    }

    try {
      if (migrateWorkspaceDataInPlace(ws, tenantRoot)) {
        await ws.save();
      }
    } catch (migErr) {
      console.error('PUT /workspace migrate:', migErr.message);
    }

    const existingNormalized = normalizeWorkspacePayload(ws.data);
    const incoming = normalizeWorkspacePayload(req.body);
    if (Array.isArray(incoming.tasks)) {
      incoming.tasks = incoming.tasks.filter((t) => !t._sharedTask);
    }

    const merged = { ...incoming };
    const isDelegatedAdmin = isAdmin && isDelegatedTenantAdmin(req, tenantRoot);
    if (isDelegatedAdmin) {
      if (Array.isArray(incoming.tasks)) {
        merged.tasks = mergeTasksForDelegatedAdmin(existingNormalized.tasks, incoming.tasks, req.user.userId);
      } else {
        merged.tasks = existingNormalized.tasks;
      }
    }

    merged.journal = mergeJournalUserSlice(
      existingNormalized.journal,
      incoming.journal,
      req.user.userId,
      tenantRoot
    );

    if (isAdmin) {
      merged.reportToOptions = sanitizeReportToOptions({ reportToOptions: incoming.reportToOptions });
    } else {
      merged.reportToOptions = existingNormalized.reportToOptions;
    }

    if (!isAdmin) {
      merged.users = await loadTenantUsers(tenantRoot, req.user.userId);
      merged.locations = existingNormalized.locations;
      merged.segregationTypes = existingNormalized.segregationTypes;
      merged.holidays = existingNormalized.holidays;
      merged.milestones = mergeArrayByCreatedBy(
        existingNormalized.milestones,
        incoming.milestones,
        req.user.userId
      );
      merged.notes = mergeArrayByCreatedBy(existingNormalized.notes, incoming.notes, req.user.userId);
      merged.learningNotes = mergeArrayByCreatedBy(
        existingNormalized.learningNotes,
        incoming.learningNotes,
        req.user.userId
      );
      merged.dailyPlanner = mergeArrayByCreatedBy(
        existingNormalized.dailyPlanner,
        incoming.dailyPlanner,
        req.user.userId
      );
      merged.codeSnippets = mergeArrayByCreatedBy(
        existingNormalized.codeSnippets,
        incoming.codeSnippets,
        req.user.userId
      );
      merged.templateBlocks = mergeArrayByCreatedBy(
        existingNormalized.templateBlocks,
        incoming.templateBlocks,
        req.user.userId
      );
    } else {
      if (!isDelegatedAdmin) {
        try {
          const previousUserIds = previousWorkspaceUserIdSet(existingNormalized.users);
          const incomingUsersMerged = await mergeIncomingUsersWithDbTenantRoster(
            incoming.users,
            tenantRoot,
            previousUserIds
          );
          await syncUsersFromClientPayload(incomingUsersMerged, { isAdmin: true, tenantRootUserId: tenantRoot });
          if (Array.isArray(incomingUsersMerged) && incomingUsersMerged.length > 0) {
            await deleteUsersNotInPayload(
              incomingUsersMerged.map((u) => u.id),
              tenantRoot,
              previousUserIds,
              false
            );
          }
        } catch (syncErr) {
          console.error('PUT /workspace: user sync error (workspace will still save):', syncErr);
        }
      }
      merged.milestones = mergeArrayByCreatedBy(
        existingNormalized.milestones,
        incoming.milestones,
        req.user.userId
      );
      merged.notes = mergeArrayByCreatedBy(existingNormalized.notes, incoming.notes, req.user.userId);
      merged.learningNotes = mergeArrayByCreatedBy(
        existingNormalized.learningNotes,
        incoming.learningNotes,
        req.user.userId
      );
      merged.dailyPlanner = mergeArrayByCreatedBy(
        existingNormalized.dailyPlanner,
        incoming.dailyPlanner,
        req.user.userId
      );
      merged.codeSnippets = mergeArrayByCreatedBy(
        existingNormalized.codeSnippets,
        incoming.codeSnippets,
        req.user.userId
      );
      merged.templateBlocks = mergeArrayByCreatedBy(
        existingNormalized.templateBlocks,
        incoming.templateBlocks,
        req.user.userId
      );
    }

    const users = await loadTenantUsers(tenantRoot, req.user.userId);
    console.log(`PUT /workspace uid=${req.user.userId} tenantRoot=${tenantRoot} users=${users.length}`);
    merged.users = users;

    ws.data = merged;
    ws.markModified('data');
    await ws.save();

    merged._debug = {
      tenantRoot,
      reqUserId: req.user.userId,
      usersReturned: users.length,
      userIds: users.map((u) => u.id),
    };

    if (ws.updatedAt) {
      merged._workspaceUpdatedAt = ws.updatedAt.toISOString();
    }

    let payload = { ...merged };
    if (!isAdmin) {
      payload = scopeWorkspaceForTenantUser(payload, req.user.userId, tenantRoot);
    } else {
      const nestedJ = ensureJournalNested(merged.journal, tenantRoot);
      payload = scopePersonalListsForUser(
        { ...merged, journal: journalFlatForUser(nestedJ, req.user.userId) },
        req.user.userId
      );
    }
    try {
      const site = await getSiteSettings();
      const siteRt = Array.isArray(site.reportToOptions) ? site.reportToOptions : [];
      const tenantSan = sanitizeReportToOptions({ reportToOptions: merged.reportToOptions });
      if (isAdmin) {
        payload.tenantReportToOptions = tenantSan;
      }
      payload.reportToOptions = tenantSan.length > 0 ? tenantSan : siteRt;
    } catch (_) {
      payload.reportToOptions = [];
    }
    return res.json(payload);
  } catch (e) {
    console.error('PUT /workspace CRASH:', e);
    return res.status(500).json({ error: 'Failed to save workspace', detail: e.message });
  }
});

router.patch('/shared-task', authMiddleware, async (req, res) => {
  try {
    const { sourceWorkspace, taskId, updates } = req.body || {};
    if (!sourceWorkspace || !taskId || !updates) {
      return res.status(400).json({ error: 'sourceWorkspace, taskId, and updates required' });
    }
    const root = Number(sourceWorkspace);
    const ws = await Workspace.findOne({ tenantRootUserId: root });
    if (!ws || !ws.data || !Array.isArray(ws.data.tasks)) {
      return res.status(404).json({ error: 'Workspace not found' });
    }
    const task = ws.data.tasks.find((t) => String(t.id) === String(taskId));
    if (!task) {
      return res.status(404).json({ error: 'Task not found' });
    }
    if (Number(task.assigned_to) !== req.user.userId) {
      return res.status(403).json({ error: 'Not assigned to you' });
    }
    const allowed = ['status', 'completion_percentage'];
    for (const key of allowed) {
      if (updates[key] !== undefined) task[key] = updates[key];
    }
    ws.markModified('data');
    await ws.save();
    return res.json({ ok: true });
  } catch (e) {
    console.error('PATCH /shared-task error:', e);
    return res.status(500).json({ error: 'Update failed' });
  }
});

// ── Reminder preference endpoints (only functional when SMTP is configured) ──

router.get('/reminder-prefs', authMiddleware, async (req, res) => {
  if (!EMAIL_CONFIGURED) return res.json({ enabled: false });
  try {
    const pref = await ReminderPreference.findOne({ userId: req.user.userId }).lean();
    return res.json({
      enabled: true,
      beforeDueDate: pref ? pref.beforeDueDate : true,
      afterDueDate: pref ? pref.afterDueDate : true,
      notifyOnAssign: pref ? pref.notifyOnAssign !== false : true,
      notifyOnSelfAssign: pref ? !!pref.notifyOnSelfAssign : false,
      setByAdmin: pref ? !!pref.setByAdmin : false,
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to load preferences' });
  }
});

router.put('/reminder-prefs', authMiddleware, async (req, res) => {
  if (!EMAIL_CONFIGURED) return res.json({ enabled: false });
  try {
    const { beforeDueDate, afterDueDate, notifyOnAssign, notifyOnSelfAssign } = req.body || {};
    const existing = await ReminderPreference.findOne({ userId: req.user.userId });
    if (existing && existing.setByAdmin && req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Your reminder preferences are managed by your admin.' });
    }
    await ReminderPreference.findOneAndUpdate(
      { userId: req.user.userId },
      {
        beforeDueDate: !!beforeDueDate,
        afterDueDate: !!afterDueDate,
        notifyOnAssign: notifyOnAssign !== false,
        notifyOnSelfAssign: !!notifyOnSelfAssign,
        setByAdmin: false,
      },
      { upsert: true, new: true }
    );
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to save preferences' });
  }
});

router.put('/reminder-prefs/:userId', authMiddleware, async (req, res) => {
  if (!EMAIL_CONFIGURED) return res.json({ enabled: false });
  if (req.user.role !== 'admin' && !req.user.isMaster) {
    return res.status(403).json({ error: 'Admin only' });
  }
  try {
    const targetUserId = parseInt(req.params.userId, 10);
    if (Number.isNaN(targetUserId)) return res.status(400).json({ error: 'Invalid userId' });
    const { beforeDueDate, afterDueDate, notifyOnAssign, notifyOnSelfAssign } = req.body || {};
    await ReminderPreference.findOneAndUpdate(
      { userId: targetUserId },
      {
        beforeDueDate: !!beforeDueDate,
        afterDueDate: !!afterDueDate,
        notifyOnAssign: notifyOnAssign !== false,
        notifyOnSelfAssign: !!notifyOnSelfAssign,
        setByAdmin: true,
      },
      { upsert: true, new: true }
    );
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to save preferences' });
  }
});

router.get('/reminder-prefs/org-users', authMiddleware, async (req, res) => {
  if (!EMAIL_CONFIGURED) return res.json({ enabled: false, users: [] });
  if (req.user.role !== 'admin' && !req.user.isMaster) {
    return res.status(403).json({ error: 'Admin only' });
  }
  try {
    const tenantRoot = resolveTenantRoot(req);
    if (tenantRoot == null) return res.json({ enabled: true, users: [] });
    const users = await User.find({
      isActive: true,
      isMaster: { $ne: true },
      $or: [{ tenantRootUserId: tenantRoot }, { userId: tenantRoot }],
    }).lean();
    const prefs = await ReminderPreference.find({
      userId: { $in: users.map((u) => u.userId) },
    }).lean();
    const prefsMap = new Map(prefs.map((p) => [p.userId, p]));
    const result = users.map((u) => {
      const p = prefsMap.get(u.userId);
      return {
        userId: u.userId,
        name: u.name,
        email: u.email,
        beforeDueDate: p ? p.beforeDueDate : true,
        afterDueDate: p ? p.afterDueDate : true,
        notifyOnAssign: p ? p.notifyOnAssign !== false : true,
        notifyOnSelfAssign: p ? !!p.notifyOnSelfAssign : false,
        setByAdmin: p ? !!p.setByAdmin : false,
      };
    });
    return res.json({ enabled: true, users: result });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to load org preferences' });
  }
});

router.post('/notify-task-assigned', authMiddleware, async (req, res) => {
  if (!EMAIL_CONFIGURED) return res.json({ ok: true, skipped: true });
  try {
    const { assignedToUserId, taskTitle, dueDate, isSelf, eventKind } = req.body || {};
    const assigneeId = parseInt(String(assignedToUserId), 10);
    if (Number.isNaN(assigneeId)) return res.status(400).json({ error: 'Invalid assignedToUserId' });

    const assignee = await User.findOne({ userId: assigneeId }).lean();
    if (!assignee || !assignee.email) return res.json({ ok: true, skipped: true });

    const tenantRoot = resolveTenantRoot(req);
    if (tenantRoot != null && !userInTenantWorkspace(tenantRoot, assignee)) {
      return res.status(403).json({ error: 'Assignee not in your organisation' });
    }

    const pref = await ReminderPreference.findOne({ userId: assigneeId }).lean();
    const wantsAssignNotify = pref ? pref.notifyOnAssign !== false : true;
    const wantsSelfNotify = pref ? !!pref.notifyOnSelfAssign : false;

    if (!wantsAssignNotify) return res.json({ ok: true, skipped: true });
    if (isSelf && !wantsSelfNotify) return res.json({ ok: true, skipped: true });

    const assigner = await User.findOne({ userId: req.user.userId }).lean();
    const assignerName = assigner ? (assigner.name || assigner.email) : 'Admin';

    const kind = eventKind === 'reassigned' ? 'reassigned' : 'created';
    await sendTaskAssignmentEmail(
      assignee.email,
      assignee.name || assignee.email,
      taskTitle || '(untitled)',
      dueDate || null,
      assignerName,
      !!isSelf,
      kind
    );
    return res.json({ ok: true });
  } catch (e) {
    console.error('notify-task-assigned error:', e);
    return res.json({ ok: true, skipped: true });
  }
});

router.post('/notify-task-rejected', authMiddleware, async (req, res) => {
  if (!EMAIL_CONFIGURED) return res.json({ ok: true, skipped: true });
  try {
    if (req.user.isMaster || req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin only' });
    }
    const { assignedToUserId, taskTitle, comment } = req.body || {};
    const assigneeId = parseInt(String(assignedToUserId), 10);
    const c = String(comment || '').trim();
    if (Number.isNaN(assigneeId) || !c) {
      return res.status(400).json({ error: 'assignedToUserId and non-empty comment required' });
    }

    const tenantRoot = resolveTenantRoot(req);
    const assignee = await User.findOne({ userId: assigneeId }).lean();
    if (!assignee || !assignee.email) return res.json({ ok: true, skipped: true });
    if (tenantRoot != null && !userInTenantWorkspace(tenantRoot, assignee)) {
      return res.status(403).json({ error: 'User not in your organisation' });
    }

    const admin = await User.findOne({ userId: req.user.userId }).lean();
    const adminName = admin ? (admin.name || admin.email) : 'Admin';

    await sendTaskRejectedEmail(
      assignee.email,
      assignee.name || assignee.email,
      taskTitle || '(untitled)',
      c,
      adminName
    );
    return res.json({ ok: true });
  } catch (e) {
    console.error('notify-task-rejected error:', e);
    return res.json({ ok: true, skipped: true });
  }
});

router.post('/notify-user-created', authMiddleware, async (req, res) => {
  if (!EMAIL_CONFIGURED) return res.json({ ok: true, skipped: true });
  try {
    if (req.user.isMaster || req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin only' });
    }
    const { newUserId, newUserEmail, newUserName } = req.body || {};
    const uid = parseInt(String(newUserId), 10);
    const em = String(newUserEmail || '').toLowerCase().trim();
    if (Number.isNaN(uid) || !em) {
      return res.status(400).json({ error: 'newUserId and newUserEmail required' });
    }

    const tenantRoot = resolveTenantRoot(req);
    const u = await User.findOne({ userId: uid }).lean();
    if (!u || u.email !== em) {
      return res.status(400).json({ error: 'User not found or email mismatch — save the user first, then retry.' });
    }
    if (tenantRoot != null && !userInTenantWorkspace(tenantRoot, u)) {
      return res.status(403).json({ error: 'User not in your organisation' });
    }

    await sendAccountCreatedEmail(u.email, newUserName || u.name || u.email, 'admin_created');
    return res.json({ ok: true });
  } catch (e) {
    console.error('notify-user-created error:', e);
    return res.json({ ok: true, skipped: true });
  }
});

router.post('/send-test-reminder', authMiddleware, async (req, res) => {
  if (!EMAIL_CONFIGURED) {
    return res.status(400).json({
      error: 'Email not configured. On Render, outbound SMTP is blocked — set RESEND_API_KEY and EMAIL_FROM (see server .env.example). Else set SMTP_EMAIL and SMTP_PASSWORD for SMTP.',
    });
  }
  try {
    const user = await User.findOne({ userId: req.user.userId }).lean();
    if (!user) return res.status(404).json({ error: 'User not found' });
    await sendTestEmail(user.email, user.name || user.email);
    return res.json({ ok: true, message: `Test email sent successfully to ${user.email}. Check your inbox (and spam folder).` });
  } catch (e) {
    console.error('send-test-reminder error:', e.message || e);
    const detail = e.message || 'Unknown error';
    return res.status(500).json({ error: `Test email failed: ${detail}` });
  }
});

router.post(
  '/email-task-view-summary',
  authMiddleware,
  emailTaskViewLimiter,
  validateBody(emailTaskViewSummarySchema),
  async (req, res) => {
    try {
      if (req.user.isMaster) {
        return res.status(403).json({ error: 'Master account cannot send tenant task emails' });
      }
      if (!isEmailEnabled()) {
        return res.status(503).json({ error: 'Email is not configured on this server.' });
      }
      const tenantRoot = resolveTenantRoot(req);
      const { recipients } = req.body;
      const sent = [];
      const errors = [];
      for (const block of recipients) {
        const uid = block.userId;
        const udoc = await User.findOne({ userId: uid });
        if (!udoc || !userInTenantWorkspace(tenantRoot, udoc)) {
          errors.push({ userId: uid, error: 'User is not in your organisation' });
          continue;
        }
        if (!udoc.email) {
          errors.push({ userId: uid, error: 'User has no email' });
          continue;
        }
        const tasks = Array.isArray(block.tasks) ? block.tasks : [];
        if (tasks.length === 0) continue;
        try {
          await sendTaskViewSummaryEmail(udoc.email, udoc.name || udoc.email, tasks);
          sent.push(uid);
        } catch (err) {
          console.error('email-task-view-summary:', err.message);
          errors.push({ userId: uid, error: err.message || 'Send failed' });
        }
      }
      return res.json({ ok: true, sent, errors });
    } catch (e) {
      console.error(e);
      return res.status(500).json({ error: 'Could not send emails' });
    }
  }
);

module.exports = router;
