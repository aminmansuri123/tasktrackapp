require('dotenv').config();
const express = require('express');
const mongoose = require('mongoose');
const helmet = require('helmet');
const cors = require('cors');
const cookieParser = require('cookie-parser');
const rateLimit = require('express-rate-limit');

const { PORT, MONGODB_URI, MASTER_EMAIL, MASTER_PASSWORD, parseOrigins, isProduction } = require('./config');
const User = require('./models/User');
const Workspace = require('./models/Workspace');
const bcrypt = require('bcryptjs');

const authRoutes = require('./routes/auth');
const workspaceRoutes = require('./routes/workspace');
const attachmentsRoutes = require('./routes/attachments');
const masterRoutes = require('./routes/master');
const { startReminderCron } = require('./services/reminderCron');

async function ensureMasterUser() {
  if (!MASTER_PASSWORD) {
    console.warn('MASTER_PASSWORD not set; master account will not be auto-provisioned.');
    return;
  }
  const email = MASTER_EMAIL.toLowerCase();
  const passwordHash = await bcrypt.hash(MASTER_PASSWORD, 12);
  let u = await User.findOne({ email });
  if (u) {
    u.passwordHash = passwordHash;
    u.isMaster = true;
    u.role = 'admin';
    u.isActive = true;
    u.tenantRootUserId = null;
    await u.save();
  } else {
    const userId = Date.now();
    await User.create({
      userId,
      email,
      name: 'Master',
      passwordHash,
      role: 'admin',
      isActive: true,
      isMaster: true,
      tenantRootUserId: null,
    });
  }
  console.log('Master user ensured for', email);
}

/** One-time style migration: shared workspace → first org admin as tenant root. */
async function migrateLegacyTenants() {
  try {
    const users = await User.find({});
    if (users.length === 0) return;

    const firstAdmin = users.find((x) => x.role === 'admin' && !x.isMaster);
    const anyNonMaster = users.find((x) => !x.isMaster);
    const rootId = firstAdmin?.userId ?? anyNonMaster?.userId;
    if (rootId == null) return;

    const workspaces = await Workspace.find({});
    const legacyWs = workspaces.find((w) => w.tenantRootUserId == null);
    if (legacyWs) {
      const canonical = await Workspace.findOne({ tenantRootUserId: rootId });
      if (canonical) {
        await Workspace.deleteOne({ _id: legacyWs._id });
        console.log(
          'Removed orphan legacy workspace; canonical workspace already exists for tenant root',
          rootId
        );
      } else {
        legacyWs.tenantRootUserId = rootId;
        await legacyWs.save();
        console.log('Assigned legacy workspace to tenant root userId', rootId);
      }
    }

    for (const u of users) {
      if (u.isMaster) {
        if (u.tenantRootUserId != null) {
          u.tenantRootUserId = null;
          await u.save();
        }
        continue;
      }
      if (u.tenantRootUserId == null) {
        u.tenantRootUserId = rootId;
        await u.save();
      }
    }
  } catch (e) {
    console.warn('Tenant migration warning:', e.message);
  }
}

async function dropStaleLegacyIndexes() {
  try {
    const col = Workspace.collection;
    const indexes = await col.indexes();
    const nameIdx = indexes.find((i) => i.name === 'name_1' && i.unique === true);
    if (nameIdx) {
      await col.dropIndex('name_1');
      console.log('Dropped stale unique index name_1 on workspaces collection');
    }
  } catch (e) {
    if (e.codeName !== 'IndexNotFound') {
      console.warn('dropStaleLegacyIndexes warning:', e.message);
    }
  }
}

async function main() {
  if (!MONGODB_URI) {
    console.error('MONGODB_URI is required');
    process.exit(1);
  }

  await mongoose.connect(MONGODB_URI);
  console.log('MongoDB connected');
  await dropStaleLegacyIndexes();
  await ensureMasterUser();
  await migrateLegacyTenants();

  const app = express();
  app.set('trust proxy', 1);

  const allowedOrigins = parseOrigins();

  app.use(
    helmet({
      crossOriginResourcePolicy: { policy: 'cross-origin' },
    })
  );

  app.use(
    cors({
      origin(origin, cb) {
        if (!origin) return cb(null, true);
        if (allowedOrigins.includes(origin)) return cb(null, true);
        return cb(null, false);
      },
      credentials: true,
    })
  );

  app.use(express.json({ limit: '15mb' }));
  app.use(cookieParser());

  const globalLimiter = rateLimit({
    windowMs: 15 * 60 * 1000,
    max: 500,
    standardHeaders: true,
    legacyHeaders: false,
  });

  const authLimiter = rateLimit({
    windowMs: 15 * 60 * 1000,
    max: 30,
    standardHeaders: true,
    legacyHeaders: false,
  });

  const masterLimiter = rateLimit({
    windowMs: 60 * 60 * 1000,
    max: 30,
    standardHeaders: true,
    legacyHeaders: false,
  });

  app.use(globalLimiter);

  app.get('/health', (_req, res) => res.json({ ok: true }));

  app.use('/api/auth', authLimiter, authRoutes);
  app.use('/api/workspace', workspaceRoutes);
  app.use('/api/attachments', attachmentsRoutes);
  app.use('/api/master', masterLimiter, masterRoutes);

  app.use((err, _req, res, _next) => {
    console.error(err);
    res.status(500).json({ error: 'Server error' });
  });

  startReminderCron();

  app.listen(PORT, () => {
    console.log(`API listening on port ${PORT}`);
    console.log('Allowed CORS origins:', allowedOrigins.join(', ') || '(none)');
    if (!isProduction) console.log('NODE_ENV is not production — cookies use non-secure mode');
  });
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
