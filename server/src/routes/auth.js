const express = require('express');
const crypto = require('crypto');
const bcrypt = require('bcryptjs');
const User = require('../models/User');
const {
  signToken,
  authMiddleware,
  resolveAuthDecoded,
  validateUserSession,
} = require('../middleware/auth');
const { ensureWorkspaceForTenantRoot } = require('../services/ensureWorkspace');
const { isProduction, EMAIL_CONFIGURED } = require('../config');
const { usersToClientShapeAll, usersToClientShapeForTenant } = require('../services/userSync');
const { assertRegistrationAllowed, getSiteSettings } = require('../services/registrationPolicy');
const { allocateUniqueUserId } = require('../services/userSync');
const { resolveTenantRootFromAdminPicker } = require('../services/tenantRoot');
const Workspace = require('../models/Workspace');
const { normalizeWorkspacePayload } = require('../services/defaultWorkspace');
const ApprovalRequest = require('../models/ApprovalRequest');
const { isEmailEnabled, sendAccountCreatedEmail, sendPasswordResetCodeEmail } = require('../services/emailService');
const { validateBody } = require('../middleware/validateBody');
const {
  loginBodySchema,
  registerBodySchema,
  forgotPasswordRequestSchema,
  forgotPasswordResetSchema,
  changePasswordSchema,
  requestApprovalSchema,
} = require('../validation/schemas');

const router = express.Router();

function cookieOptions() {
  return {
    httpOnly: true,
    secure: isProduction,
    sameSite: isProduction ? 'none' : 'lax',
    path: '/',
  };
}

function publicUser(u) {
  return {
    id: u.userId,
    email: u.email,
    name: u.name,
    role: u.role,
    is_active: u.isActive,
    isMaster: u.isMaster,
    enabledFeatures: Array.isArray(u.enabledFeatures) ? u.enabledFeatures : [],
  };
}

router.get('/registration-policy', async (_req, res) => {
  try {
    const s = await getSiteSettings();
    return res.json({ registrationMode: s.registrationMode || 'open' });
  } catch (e) {
    console.error(e);
    return res.json({ registrationMode: 'open' });
  }
});

router.post('/request-approval', validateBody(requestApprovalSchema), async (req, res) => {
  try {
    const { name, email } = req.body;
    const em = String(email || '').toLowerCase().trim();
    const nm = String(name || '').trim();
    if (!em || !nm) {
      return res.status(400).json({ error: 'Name and email required' });
    }
    const existing = await ApprovalRequest.findOne({ email: em, status: 'pending' });
    if (existing) {
      return res.json({ ok: true, message: 'Approval request already submitted. Please wait for master to review.' });
    }
    const already = await User.findOne({ email: em });
    if (already) {
      return res.status(400).json({ error: 'This email is already registered. Try signing in.' });
    }
    await ApprovalRequest.create({ email: em, name: nm });
    return res.json({ ok: true, message: 'Approval request submitted. You will be able to register once the master approves your request.' });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Could not submit approval request' });
  }
});

router.get('/check-email', async (req, res) => {
  try {
    const email = String(req.query.email || '').toLowerCase().trim();
    if (!email) return res.json({ exists: false });
    const doc = await User.findOne({ email }).lean();
    return res.json({ exists: !!doc });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Check failed' });
  }
});

/** Active non-master admins (for team signup: pick any admin in the org). */
router.get('/org-admins', async (_req, res) => {
  try {
    const list = await User.find({ role: 'admin', isMaster: { $ne: true }, isActive: true })
      .sort({ name: 1 })
      .lean();
    return res.json(
      list.map((u) => ({
        id: u.userId,
        name: u.name || '',
        email: u.email || '',
      }))
    );
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to load account admins' });
  }
});

router.post('/register', validateBody(registerBodySchema), async (req, res) => {
  try {
    const { name, email, password } = req.body;
    const em = String(email).toLowerCase().trim();
    const exists = await User.findOne({ email: em });
    if (exists) {
      return res.status(400).json({ error: 'Email already registered' });
    }
    const denied = await assertRegistrationAllowed(em);
    if (denied === 'APPROVAL_REQUIRED') {
      return res.status(403).json({ error: 'APPROVAL_REQUIRED' });
    }
    if (denied) {
      return res.status(403).json({ error: denied });
    }
    const accountType = req.body.accountType === 'team_user' ? 'team_user' : 'org_admin';
    const passwordHash = await bcrypt.hash(String(password), 12);
    const userId = await allocateUniqueUserId(Date.now());

    let doc;
    if (accountType === 'team_user') {
      const raw = req.body.orgAdminUserId;
      const pickerId = typeof raw === 'number' ? raw : parseInt(String(raw), 10);
      if (Number.isNaN(pickerId)) {
        return res.status(400).json({ error: 'Select an account admin' });
      }
      const tenantRoot = await resolveTenantRootFromAdminPicker(pickerId);
      if (tenantRoot == null) {
        return res.status(400).json({ error: 'Invalid account admin' });
      }
      const regNow = new Date();
      doc = await User.create({
        userId,
        email: em,
        name: String(name).trim(),
        passwordHash,
        role: 'user',
        isActive: true,
        isMaster: false,
        tenantRootUserId: tenantRoot,
        lastActivityAt: regNow,
        lastLoginAt: regNow,
      });
      await ensureWorkspaceForTenantRoot(tenantRoot);
      try {
        const wsU = await Workspace.findOne({ tenantRootUserId: tenantRoot });
        if (wsU) {
          const norm = normalizeWorkspacePayload(wsU.data);
          norm.users = await usersToClientShapeForTenant(tenantRoot);
          wsU.data = norm;
          wsU.markModified('data');
          await wsU.save();
        }
      } catch (seedErr) {
        console.error('Register: seed ws users (team_user):', seedErr.message);
      }
    } else {
      const regNow = new Date();
      doc = await User.create({
        userId,
        email: em,
        name: String(name).trim(),
        passwordHash,
        role: 'admin',
        isActive: true,
        isMaster: false,
        tenantRootUserId: userId,
        lastActivityAt: regNow,
        lastLoginAt: regNow,
      });
      await ensureWorkspaceForTenantRoot(userId);
      try {
        const wsA = await Workspace.findOne({ tenantRootUserId: userId });
        if (wsA) {
          const norm = normalizeWorkspacePayload(wsA.data);
          norm.users = await usersToClientShapeForTenant(userId);
          wsA.data = norm;
          wsA.markModified('data');
          await wsA.save();
        }
      } catch (seedErr) {
        console.error('Register: seed ws users (org_admin):', seedErr.message);
      }
    }

    const token = signToken({
      sub: doc.userId,
      role: doc.role,
      isMaster: doc.isMaster,
    });
    res.cookie('auth_token', token, cookieOptions());
    if (isEmailEnabled()) {
      void sendAccountCreatedEmail(doc.email, doc.name || doc.email, 'self_register').catch((e) =>
        console.error('Register welcome email:', e.message)
      );
    }
    const sReg = await getSiteSettings();
    const sessionIdleTimeoutMinutesReg =
      sReg.sessionIdleTimeoutMinutes != null && !Number.isNaN(Number(sReg.sessionIdleTimeoutMinutes))
        ? Math.max(0, Number(sReg.sessionIdleTimeoutMinutes))
        : 0;
    return res.status(201).json({
      user: publicUser(doc),
      token,
      smtpConfigured: EMAIL_CONFIGURED,
      sessionIdleTimeoutMinutes: sessionIdleTimeoutMinutesReg,
    });
  } catch (e) {
    console.error(e);
    if (e && e.code === 11000) {
      const dupKey = e.keyPattern && Object.keys(e.keyPattern)[0];
      if (dupKey === 'email') {
        return res.status(400).json({ error: 'Email already registered' });
      }
      return res.status(409).json({ error: 'Account setup conflict — try signing in' });
    }
    return res.status(500).json({ error: 'Registration failed' });
  }
});

const FORGOT_CODE_TTL_MS = 15 * 60 * 1000;
const FORGOT_RESEND_COOLDOWN_MS = 60 * 1000;
const FORGOT_MAX_VERIFY_ATTEMPTS = 5;

function clearPasswordResetState(doc) {
  doc.passwordResetCodeHash = null;
  doc.passwordResetExpiresAt = null;
  doc.passwordResetAttempts = 0;
  doc.passwordResetLastSentAt = null;
}

/** Same response whether or not the account exists (avoid email enumeration). */
function forgotRequestGenericResponse(res) {
  return res.json({
    ok: true,
    message: 'If an account exists for this email, a reset code will arrive shortly.',
  });
}

router.post('/forgot-password/request', validateBody(forgotPasswordRequestSchema), async (req, res) => {
  try {
    const em = String(req.body.email || '').toLowerCase().trim();
    if (!em) {
      return res.status(400).json({ error: 'Email is required' });
    }
    const doc = await User.findOne({ email: em });
    if (!doc || !doc.isActive) {
      return forgotRequestGenericResponse(res);
    }
    const now = Date.now();
    if (doc.passwordResetLastSentAt) {
      const elapsed = now - new Date(doc.passwordResetLastSentAt).getTime();
      if (elapsed < FORGOT_RESEND_COOLDOWN_MS) {
        const waitSec = Math.ceil((FORGOT_RESEND_COOLDOWN_MS - elapsed) / 1000);
        return res.status(429).json({ error: `Please wait ${waitSec}s before requesting another code.` });
      }
    }
    if (!isEmailEnabled()) {
      return res.status(503).json({ error: 'Password reset email is not configured on this server.' });
    }
    const code = String(crypto.randomInt(0, 10000)).padStart(4, '0');
    doc.passwordResetCodeHash = await bcrypt.hash(code, 10);
    doc.passwordResetExpiresAt = new Date(now + FORGOT_CODE_TTL_MS);
    doc.passwordResetAttempts = 0;
    doc.passwordResetLastSentAt = new Date(now);
    await doc.save();
    const sent = await sendPasswordResetCodeEmail(doc.email, doc.name || doc.email, code);
    if (!sent) {
      clearPasswordResetState(doc);
      await doc.save();
      return res.status(500).json({ error: 'Could not send reset email. Try again later.' });
    }
    return forgotRequestGenericResponse(res);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Could not process reset request' });
  }
});

router.post('/forgot-password/reset', validateBody(forgotPasswordResetSchema), async (req, res) => {
  try {
    const { email, code, newPassword } = req.body;
    const em = String(email || '').toLowerCase().trim();
    const codeDigits = String(code || '').replace(/\D/g, '').slice(0, 4);
    if (!em || codeDigits.length !== 4 || !newPassword || String(newPassword).length < 6) {
      return res.status(400).json({ error: 'Email, 4-digit code, and new password (min 6 characters) are required' });
    }
    const doc = await User.findOne({ email: em });
    if (!doc || !doc.isActive || !doc.passwordResetCodeHash || !doc.passwordResetExpiresAt) {
      return res.status(400).json({ error: 'Invalid or expired code. Request a new code from the sign-in page.' });
    }
    if (new Date() > doc.passwordResetExpiresAt) {
      clearPasswordResetState(doc);
      await doc.save();
      return res.status(400).json({ error: 'Code expired. Request a new one.' });
    }
    if ((doc.passwordResetAttempts || 0) >= FORGOT_MAX_VERIFY_ATTEMPTS) {
      clearPasswordResetState(doc);
      await doc.save();
      return res.status(400).json({ error: 'Too many attempts. Request a new code.' });
    }
    const match = await bcrypt.compare(codeDigits, doc.passwordResetCodeHash);
    if (!match) {
      doc.passwordResetAttempts = (doc.passwordResetAttempts || 0) + 1;
      await doc.save();
      return res.status(400).json({ error: 'Invalid code.' });
    }
    doc.passwordHash = await bcrypt.hash(String(newPassword), 12);
    clearPasswordResetState(doc);
    doc.loginLocked = false;
    doc.failedLoginAttempts = 0;
    await doc.save();
    return res.json({ ok: true, message: 'Password updated. You can sign in now.' });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Could not reset password' });
  }
});

router.post('/login', validateBody(loginBodySchema), async (req, res) => {
  try {
    const { email, password } = req.body;
    const em = String(email).toLowerCase().trim();
    const doc = await User.findOne({ email: em });
    if (!doc) {
      return res.status(401).json({ error: 'Invalid email or password' });
    }
    if (doc.loginLocked) {
      return res.status(403).json({
        error:
          'Account locked after too many failed sign-in attempts. Use Forgot password to reset and unlock.',
        code: 'LOGIN_LOCKED',
      });
    }
    if (!doc.isActive) {
      return res.status(403).json({ error: 'Account is disabled' });
    }
    const passwordOk = await bcrypt.compare(String(password), doc.passwordHash);
    if (!passwordOk) {
      doc.failedLoginAttempts = (doc.failedLoginAttempts || 0) + 1;
      if (doc.failedLoginAttempts >= 5) {
        doc.loginLocked = true;
      }
      await doc.save();
      return res.status(401).json({ error: 'Invalid email or password' });
    }
    doc.failedLoginAttempts = 0;
    doc.loginLocked = false;
    const now = new Date();
    doc.lastLoginAt = now;
    doc.lastActivityAt = now;
    await doc.save();
    const token = signToken({
      sub: doc.userId,
      role: doc.role,
      isMaster: doc.isMaster,
    });
    res.cookie('auth_token', token, cookieOptions());
    const s = await getSiteSettings();
    const sessionIdleTimeoutMinutes =
      s.sessionIdleTimeoutMinutes != null && !Number.isNaN(Number(s.sessionIdleTimeoutMinutes))
        ? Math.max(0, Number(s.sessionIdleTimeoutMinutes))
        : 0;
    return res.json({
      user: publicUser(doc),
      token,
      smtpConfigured: EMAIL_CONFIGURED,
      sessionIdleTimeoutMinutes,
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Login failed' });
  }
});

router.post('/logout', (_req, res) => {
  res.clearCookie('auth_token', { path: '/', sameSite: isProduction ? 'none' : 'lax', secure: isProduction });
  res.json({ ok: true });
});

router.post('/change-password', authMiddleware, validateBody(changePasswordSchema), async (req, res) => {
  try {
    const { currentPassword, newPassword } = req.body;
    const doc = await User.findOne({ userId: req.user.userId });
    if (!doc || !(await bcrypt.compare(String(currentPassword), doc.passwordHash))) {
      return res.status(401).json({ error: 'Current password is incorrect' });
    }
    doc.passwordHash = await bcrypt.hash(String(newPassword), 12);
    doc.lastActivityAt = new Date();
    await doc.save();
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Could not update password' });
  }
});

router.get('/me', async (req, res) => {
  const decoded = resolveAuthDecoded(req);
  if (!decoded) {
    return res.status(401).json({ error: 'Not logged in' });
  }
  const userId = typeof decoded.sub === 'number' ? decoded.sub : parseInt(String(decoded.sub), 10);
  if (Number.isNaN(userId)) {
    return res.status(401).json({ error: 'Invalid session' });
  }
  const doc = await User.findOne({ userId });
  const sess = await validateUserSession(doc, res);
  if (!sess.ok) {
    return;
  }
  return res.json({
    ...publicUser(doc),
    smtpConfigured: EMAIL_CONFIGURED,
    sessionIdleTimeoutMinutes: sess.sessionIdleTimeoutMinutes,
  });
});

router.get('/users-for-master', async (req, res) => {
  const decoded = resolveAuthDecoded(req);
  if (!decoded) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  if (!decoded.isMaster) {
    return res.status(403).json({ error: 'Master only' });
  }
  const list = await usersToClientShapeAll();
  return res.json(list);
});

module.exports = router;
