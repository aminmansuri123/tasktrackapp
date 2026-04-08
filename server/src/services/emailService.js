const dns = require('dns');
const nodemailer = require('nodemailer');
const {
  SMTP_EMAIL,
  SMTP_PASSWORD,
  SMTP_HOST,
  SMTP_PORT,
  SMTP_CONFIGURED,
  RESEND_API_KEY,
  EMAIL_FROM,
  EMAIL_CONFIGURED,
} = require('../config');
const { getSiteSettings } = require('./registrationPolicy');

function isEmailEnabled() {
  return EMAIL_CONFIGURED;
}

function resendFromHeader() {
  if (EMAIL_FROM) return EMAIL_FROM;
  return 'Task Tracker <onboarding@resend.dev>';
}

function normalizeCcList(cc) {
  if (!Array.isArray(cc)) return [];
  return [...new Set(cc.map((e) => String(e).trim().toLowerCase()).filter((e) => e.includes('@')))];
}

async function getStatusEmailCcList() {
  try {
    const s = await getSiteSettings();
    return normalizeCcList(s.statusEmailCc);
  } catch {
    return [];
  }
}

async function getEmailTemplatePair(key) {
  try {
    const s = await getSiteSettings();
    const t = s.emailTemplates && s.emailTemplates[key];
    if (!t || typeof t !== 'object') return { subject: '', bodyHtml: '' };
    return {
      subject: String(t.subject || '').trim(),
      bodyHtml: String(t.bodyHtml || '').trim(),
    };
  } catch {
    return { subject: '', bodyHtml: '' };
  }
}

function interpolateEmailTemplate(str, vars) {
  if (!str) return '';
  let out = String(str);
  for (const [k, v] of Object.entries(vars)) {
    const re = new RegExp(`\\{\\{\\s*${k}\\s*\\}\\}`, 'g');
    out = out.replace(re, v != null ? String(v) : '');
  }
  return out;
}

async function sendMailResend(toEmail, subject, html, ccList = []) {
  const payload = {
    from: resendFromHeader(),
    to: [toEmail],
    subject,
    html,
  };
  const cc = normalizeCcList(ccList);
  if (cc.length) payload.cc = cc;
  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${RESEND_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });
  const bodyText = await res.text();
  if (!res.ok) {
    throw new Error(`Resend API ${res.status}: ${bodyText.slice(0, 500)}`);
  }
}

async function sendMailSmtp(toEmail, subject, html, ccList = []) {
  const transporter = await getTransporter();
  if (!transporter) throw new Error('SMTP transporter unavailable');
  const cc = normalizeCcList(ccList);
  const mail = {
    from: `"Task Tracker" <${SMTP_EMAIL}>`,
    to: toEmail,
    subject,
    html,
  };
  if (cc.length) mail.cc = cc.join(',');
  await transporter.sendMail(mail);
}

let _loggedTransport = false;
async function sendMail(toEmail, subject, html, options = {}) {
  const toLower = String(toEmail || '').trim().toLowerCase();
  const extraCc = normalizeCcList(options.cc);
  const statusCc = options.skipStatusCc ? [] : await getStatusEmailCcList();
  const ccList = normalizeCcList([...statusCc, ...extraCc]).filter((e) => e && e !== toLower);

  if (RESEND_API_KEY) {
    if (!_loggedTransport) {
      console.log('[email] Sending via Resend API (HTTPS)');
      _loggedTransport = true;
    }
    await sendMailResend(toEmail, subject, html, ccList);
    return;
  }
  if (!_loggedTransport) {
    console.log('[email] Sending via SMTP');
    _loggedTransport = true;
  }
  await sendMailSmtp(toEmail, subject, html, ccList);
}

let _transporter = null;
let _resolvedIpv4 = null;

function resolveHostIpv4() {
  return new Promise((resolve) => {
    dns.lookup(SMTP_HOST, { family: 4 }, (err, address) => {
      if (err) {
        console.warn('[email] IPv4 DNS lookup failed, using hostname:', err.message);
        resolve(SMTP_HOST);
      } else {
        console.log(`[email] Resolved ${SMTP_HOST} → ${address} (IPv4)`);
        resolve(address);
      }
    });
  });
}

async function getTransporter() {
  if (!SMTP_CONFIGURED) return null;
  if (_transporter) return _transporter;

  if (!_resolvedIpv4) {
    _resolvedIpv4 = await resolveHostIpv4();
  }

  _transporter = nodemailer.createTransport({
    host: _resolvedIpv4,
    port: SMTP_PORT,
    secure: SMTP_PORT === 465,
    auth: { user: SMTP_EMAIL, pass: SMTP_PASSWORD },
    tls: { servername: SMTP_HOST },
    connectionTimeout: 10000,
    greetingTimeout: 10000,
    socketTimeout: 15000,
  });
  console.log(`[email] Transporter created: ${_resolvedIpv4}:${SMTP_PORT} (from ${SMTP_HOST}), secure=${SMTP_PORT === 465}`);
  return _transporter;
}

function escapeHtml(s) {
  if (s == null || s === '') return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Appended to every HTML email body for consistent branding. */
function emailBrandFooterHtml() {
  return '<p style="font-size:12px;color:#94a3b8;margin:20px 0 0;padding-top:16px;border-top:1px solid #e2e8f0;text-align:center;line-height:1.5;">Live — developed by Amin Mansuri</p>';
}

function formatDate(dateStr) {
  if (!dateStr) return 'N/A';
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('en-IN', { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' });
}

function buildTaskRow(task, type) {
  const color = type === 'overdue' ? '#e74c3c' : '#f39c12';
  const label = type === 'overdue' ? 'Overdue' : 'Due Tomorrow';
  const title = task.title || task.task_name || '(untitled)';
  const due = formatDate(task.due_date || task.next_due_date);
  const status = task.status || 'pending';
  return `
    <tr>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;">${title}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;">${due}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;">${status}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;"><span style="color:${color};font-weight:600;">${label}</span></td>
    </tr>`;
}

function buildReminderHtml(userName, overdueTasks, upcomingTasks) {
  const overdueRows = overdueTasks.map((t) => buildTaskRow(t, 'overdue')).join('');
  const upcomingRows = upcomingTasks.map((t) => buildTaskRow(t, 'upcoming')).join('');
  const totalCount = overdueTasks.length + upcomingTasks.length;

  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;color:#333;margin:0;padding:20px;background:#f5f5f5;">
  <div style="max-width:640px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">
    <div style="background:#2563eb;color:#fff;padding:20px 24px;">
      <h1 style="margin:0;font-size:20px;">Task Reminder</h1>
      <p style="margin:6px 0 0;font-size:14px;opacity:0.9;">Hello ${userName}, you have ${totalCount} task${totalCount !== 1 ? 's' : ''} that need attention.</p>
    </div>
    <div style="padding:20px 24px;">
      ${overdueTasks.length > 0 ? `
      <h2 style="font-size:16px;color:#e74c3c;margin:0 0 10px;">Overdue Tasks (${overdueTasks.length})</h2>
      <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">
        <thead><tr style="background:#fafafa;text-align:left;">
          <th style="padding:8px 12px;border-bottom:2px solid #eee;">Task</th>
          <th style="padding:8px 12px;border-bottom:2px solid #eee;">Due Date</th>
          <th style="padding:8px 12px;border-bottom:2px solid #eee;">Status</th>
          <th style="padding:8px 12px;border-bottom:2px solid #eee;">Type</th>
        </tr></thead>
        <tbody>${overdueRows}</tbody>
      </table>` : ''}
      ${upcomingTasks.length > 0 ? `
      <h2 style="font-size:16px;color:#f39c12;margin:0 0 10px;">Due Tomorrow (${upcomingTasks.length})</h2>
      <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">
        <thead><tr style="background:#fafafa;text-align:left;">
          <th style="padding:8px 12px;border-bottom:2px solid #eee;">Task</th>
          <th style="padding:8px 12px;border-bottom:2px solid #eee;">Due Date</th>
          <th style="padding:8px 12px;border-bottom:2px solid #eee;">Status</th>
          <th style="padding:8px 12px;border-bottom:2px solid #eee;">Type</th>
        </tr></thead>
        <tbody>${upcomingRows}</tbody>
      </table>` : ''}
      <p style="font-size:13px;color:#999;margin-top:16px;">This is an automated reminder from your Task Management System.</p>
      ${emailBrandFooterHtml()}
    </div>
  </div>
</body>
</html>`;
}

async function sendTaskReminderEmail(toEmail, userName, overdueTasks, upcomingTasks) {
  if (!isEmailEnabled()) return false;

  const total = overdueTasks.length + upcomingTasks.length;
  const defaultSubject =
    total === 1 ? 'Task Reminder: 1 task needs attention' : `Task Reminder: ${total} tasks need attention`;
  const tpl = await getEmailTemplatePair('reminder');
  const defaultBody = buildReminderHtml(userName, overdueTasks, upcomingTasks);
  const plain = { userName: userName || '', totalCount: String(total) };
  const htmlVars = {
    ...plain,
    userName: escapeHtml(plain.userName),
    defaultBody,
  };
  const subject = tpl.subject ? interpolateEmailTemplate(tpl.subject, plain) : defaultSubject;
  const html = tpl.bodyHtml ? interpolateEmailTemplate(tpl.bodyHtml, htmlVars) : defaultBody;

  try {
    await sendMail(toEmail, subject, html);
    return true;
  } catch (err) {
    console.error(`Email send failed for ${toEmail}:`, err.message);
    return false;
  }
}

function buildAssignmentHtml(assigneeName, taskTitle, dueDate, assignerName, isSelf) {
  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;color:#333;margin:0;padding:20px;background:#f5f5f5;">
  <div style="max-width:640px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">
    <div style="background:#16a34a;color:#fff;padding:20px 24px;">
      <h1 style="margin:0;font-size:20px;">New Task Assigned</h1>
      <p style="margin:6px 0 0;font-size:14px;opacity:0.9;">Hello ${escapeHtml(assigneeName)}, a task has been ${isSelf ? 'created by you' : 'assigned to you'}.</p>
    </div>
    <div style="padding:20px 24px;">
      <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:16px;">
        <tr><td style="padding:8px 0;font-weight:600;width:100px;">Task:</td><td style="padding:8px 0;">${escapeHtml(taskTitle || '(untitled)')}</td></tr>
        <tr><td style="padding:8px 0;font-weight:600;">Due Date:</td><td style="padding:8px 0;">${dueDate ? formatDate(dueDate) : 'Not set'}</td></tr>
        <tr><td style="padding:8px 0;font-weight:600;">Assigned By:</td><td style="padding:8px 0;">${isSelf ? 'Self' : escapeHtml(assignerName || 'Admin')}</td></tr>
      </table>
      <p style="font-size:13px;color:#666;">${isSelf ? 'You created this task for yourself.' : `Assigned by <strong>${escapeHtml(assignerName || 'an admin')}</strong>.`}</p>
      <p style="font-size:13px;color:#999;margin-top:16px;">This is an automated notification from your Task Management System.</p>
      ${emailBrandFooterHtml()}
    </div>
  </div>
</body>
</html>`;
}

async function sendTaskAssignmentEmail(toEmail, assigneeName, taskTitle, dueDate, assignerName, isSelf, eventKind) {
  if (!isEmailEnabled()) return false;

  const title = taskTitle || '(untitled)';
  let defaultSubject;
  if (eventKind === 'reassigned' && !isSelf) {
    defaultSubject = `Task reassigned to you: ${title}`;
  } else if (isSelf) {
    defaultSubject = `New task created: ${title}`;
  } else {
    defaultSubject = `New task assigned to you: ${title}`;
  }

  const tpl = await getEmailTemplatePair('task_assigned');
  const dueFmt = dueDate ? formatDate(dueDate) : 'Not set';
  const plain = {
    assigneeName: assigneeName || '',
    taskTitle: title,
    dueDateFormatted: dueFmt,
    assignerName: assignerName || 'Admin',
    isSelf: isSelf ? 'yes' : 'no',
    eventKind: eventKind || '',
  };
  const htmlVars = {
    ...plain,
    assigneeName: escapeHtml(plain.assigneeName),
    taskTitle: escapeHtml(plain.taskTitle),
    assignerName: escapeHtml(plain.assignerName),
    isSelfText: isSelf
      ? 'You created this task for yourself.'
      : `Assigned by <strong>${escapeHtml(assignerName || 'an admin')}</strong>.`,
    defaultBody: buildAssignmentHtml(assigneeName, taskTitle, dueDate, assignerName, isSelf),
  };
  const subject = tpl.subject ? interpolateEmailTemplate(tpl.subject, plain) : defaultSubject;
  const html = tpl.bodyHtml ? interpolateEmailTemplate(tpl.bodyHtml, htmlVars) : htmlVars.defaultBody;

  try {
    await sendMail(toEmail, subject, html);
    return true;
  } catch (err) {
    console.error(`Assignment email failed for ${toEmail}:`, err.message);
    return false;
  }
}

function buildTaskRejectedHtml(assigneeName, taskTitle, comment, adminName) {
  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;color:#333;margin:0;padding:20px;background:#f5f5f5;">
  <div style="max-width:640px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">
    <div style="background:#dc2626;color:#fff;padding:20px 24px;">
      <h1 style="margin:0;font-size:20px;">Task completion not accepted</h1>
      <p style="margin:6px 0 0;font-size:14px;opacity:0.95;">Hello ${escapeHtml(assigneeName)}, your submitted completion for a task was reviewed and needs follow-up.</p>
    </div>
    <div style="padding:20px 24px;">
      <p style="margin:0 0 12px;"><strong>Task:</strong> ${escapeHtml(taskTitle || '(untitled)')}</p>
      <p style="margin:0 0 8px;"><strong>Reviewer:</strong> ${escapeHtml(adminName || 'Admin')}</p>
      <div style="background:#fef2f2;border-left:4px solid #dc2626;padding:12px 14px;margin-top:16px;border-radius:4px;">
        <p style="margin:0 0 6px;font-size:13px;color:#991b1b;font-weight:600;">Comment</p>
        <p style="margin:0;white-space:pre-wrap;font-size:14px;color:#333;">${escapeHtml(comment)}</p>
      </div>
      <p style="font-size:13px;color:#999;margin-top:20px;">Please update the task as needed and resubmit when ready.</p>
      ${emailBrandFooterHtml()}
    </div>
  </div>
</body>
</html>`;
}

async function sendTaskRejectedEmail(toEmail, assigneeName, taskTitle, comment, adminName) {
  if (!isEmailEnabled()) return false;
  const title = taskTitle || '(untitled)';
  const defaultSubject = `Task needs revision: ${title}`;
  const tpl = await getEmailTemplatePair('task_rejected');
  const plain = {
    assigneeName: assigneeName || '',
    taskTitle: title,
    adminComment: comment || '',
    adminName: adminName || 'Admin',
  };
  const htmlVars = {
    ...plain,
    assigneeName: escapeHtml(plain.assigneeName),
    taskTitle: escapeHtml(plain.taskTitle),
    adminComment: escapeHtml(plain.adminComment),
    adminName: escapeHtml(plain.adminName),
    defaultBody: buildTaskRejectedHtml(assigneeName, taskTitle, comment, adminName),
  };
  const subject = tpl.subject ? interpolateEmailTemplate(tpl.subject, plain) : defaultSubject;
  const html = tpl.bodyHtml ? interpolateEmailTemplate(tpl.bodyHtml, htmlVars) : htmlVars.defaultBody;
  try {
    await sendMail(toEmail, subject, html);
    return true;
  } catch (err) {
    console.error(`Task rejected email failed for ${toEmail}:`, err.message);
    return false;
  }
}

function buildAccountCreatedHtml(userName, contextHtml) {
  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;color:#333;margin:0;padding:20px;background:#f5f5f5;">
  <div style="max-width:640px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">
    <div style="background:#2563eb;color:#fff;padding:20px 24px;">
      <h1 style="margin:0;font-size:20px;">Welcome to Task Tracker</h1>
      <p style="margin:6px 0 0;font-size:14px;opacity:0.9;">Hello ${escapeHtml(userName)},</p>
    </div>
    <div style="padding:20px 24px;">
      ${contextHtml}
      <p style="font-size:14px;color:#333;margin-top:16px;">Sign in with <strong>this email address</strong> and your password.</p>
      <p style="font-size:13px;color:#999;margin-top:16px;">This message was sent automatically. Do not reply with your password.</p>
      ${emailBrandFooterHtml()}
    </div>
  </div>
</body>
</html>`;
}

async function sendAccountCreatedEmail(toEmail, userName, source) {
  if (!isEmailEnabled()) return false;
  let contextHtml;
  if (source === 'self_register') {
    contextHtml = '<p style="font-size:14px;color:#333;margin:0;">Your account was created successfully. Use the password you chose at registration.</p>';
  } else {
    contextHtml =
      '<p style="font-size:14px;color:#333;margin:0;">An administrator created your Task Tracker account. Use the password they gave you (you can change it after signing in).</p>';
  }
  const defaultSubject = 'Your Task Tracker account is ready';
  const tpl = await getEmailTemplatePair('account_created');
  const plain = { userName: userName || '', source: source || '' };
  const htmlVars = {
    ...plain,
    userName: escapeHtml(plain.userName),
    contextHtml,
    defaultBody: buildAccountCreatedHtml(userName, contextHtml),
  };
  const subject = tpl.subject ? interpolateEmailTemplate(tpl.subject, plain) : defaultSubject;
  const html = tpl.bodyHtml ? interpolateEmailTemplate(tpl.bodyHtml, htmlVars) : htmlVars.defaultBody;
  try {
    await sendMail(toEmail, subject, html);
    return true;
  } catch (err) {
    console.error(`Account created email failed for ${toEmail}:`, err.message);
    return false;
  }
}

function buildPasswordResetCodeHtml(userName, code) {
  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;color:#333;margin:0;padding:20px;background:#f5f5f5;">
  <div style="max-width:640px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">
    <div style="background:#0d47a1;color:#fff;padding:20px 24px;">
      <h1 style="margin:0;font-size:20px;">Password reset code</h1>
      <p style="margin:6px 0 0;font-size:14px;opacity:0.95;">Hello ${escapeHtml(userName)},</p>
    </div>
    <div style="padding:20px 24px;">
      <p style="font-size:14px;color:#333;margin:0 0 16px;">Use this code to reset your Task Management System password. It expires in 15 minutes.</p>
      <p style="font-size:32px;font-weight:700;letter-spacing:0.25em;text-align:center;margin:20px 0;padding:16px;background:#f1f5f9;border-radius:8px;color:#0d47a1;">${escapeHtml(code)}</p>
      <p style="font-size:13px;color:#666;margin:0;">If you did not request this, you can ignore this email.</p>
      ${emailBrandFooterHtml()}
    </div>
  </div>
</body>
</html>`;
}

async function sendPasswordResetCodeEmail(toEmail, userName, code) {
  if (!isEmailEnabled()) return false;
  const defaultSubject = `Your password reset code: ${code}`;
  const tpl = await getEmailTemplatePair('password_reset');
  const plain = { userName: userName || '', code: code || '' };
  const htmlVars = {
    ...plain,
    userName: escapeHtml(plain.userName),
    code: escapeHtml(plain.code),
    defaultBody: buildPasswordResetCodeHtml(userName, code),
  };
  const subject = tpl.subject ? interpolateEmailTemplate(tpl.subject, plain) : defaultSubject;
  const html = tpl.bodyHtml ? interpolateEmailTemplate(tpl.bodyHtml, htmlVars) : htmlVars.defaultBody;
  try {
    await sendMail(toEmail, subject, html);
    return true;
  } catch (err) {
    console.error(`Password reset email failed for ${toEmail}:`, err.message);
    return false;
  }
}

async function sendTestEmail(toEmail, userName) {
  if (!isEmailEnabled()) {
    throw new Error('Email not configured. Set RESEND_API_KEY (recommended on Render) or SMTP_EMAIL + SMTP_PASSWORD.');
  }
  const demoTask = {
    title: 'Sample Task — Test Reminder',
    due_date: new Date().toISOString().split('T')[0],
    status: 'in_progress',
  };
  await sendMail(toEmail, 'Test Email — Task Tracker Reminder', buildReminderHtml(userName, [], [demoTask]));
  return true;
}

/** Dashboard-aligned status cell colors (recurring report / Task View). */
function taskViewEmailStatusStyle(statusLabel) {
  const s = String(statusLabel || '').trim();
  const map = {
    Completed: { bg: '#b7e1cd', color: '#000000' },
    'Needs Improvement': { bg: '#ffe599', color: '#000000' },
    'Need Improvement': { bg: '#ffe599', color: '#000000' },
    'In Process': { bg: '#d1ecf1', color: '#000000' },
    'Not Done': { bg: '#ff4d4f', color: '#ffffff' },
    Overdue: { bg: '#f8d7da', color: '#000000' },
    Pending: { bg: '#fff3cd', color: '#000000' },
    'No Due Date': { bg: '#e3f2fd', color: '#000000' },
  };
  return map[s] || { bg: '#ffffff', color: '#000000' };
}

function formatIstGeneratedLine() {
  try {
    const fmt = new Intl.DateTimeFormat('en-IN', {
      timeZone: 'Asia/Kolkata',
      dateStyle: 'medium',
      timeStyle: 'short',
    });
    return fmt.format(new Date());
  } catch (_) {
    return new Date().toISOString();
  }
}

/**
 * Task View “email filtered list” — one message per assignee with a simple HTML table.
 * @param {string} toEmail
 * @param {string} userName
 * @param {{ title: string, due?: string, overdue?: boolean, status?: string }[]} tasks
 */
function buildTaskViewSummaryTableRows(tasks) {
  return tasks
    .map((t) => {
      const statusRaw =
        t.status != null && String(t.status).trim() !== ''
          ? String(t.status).trim()
          : t.overdue
            ? 'Overdue'
            : 'Pending';
      const st = taskViewEmailStatusStyle(statusRaw);
      return `<tr>
<td style="padding:10px 12px;border-bottom:1px solid #bbdefb;vertical-align:middle;">${escapeHtml(t.title || '(untitled)')}</td>
<td style="padding:10px 12px;border-bottom:1px solid #bbdefb;vertical-align:middle;color:#1565c0;">${escapeHtml(t.due || '—')}</td>
<td style="padding:10px 12px;border-bottom:1px solid #bbdefb;vertical-align:middle;font-weight:600;background:${st.bg};color:${st.color};">${escapeHtml(statusRaw)}</td>
</tr>`;
    })
    .join('');
}

function buildTaskViewSummaryEmailHtml(userName, rowsHtml, istLine) {
  return `
<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#e8f4fc;">
  <div style="font-family:Segoe UI,Roboto,Helvetica,sans-serif;max-width:720px;margin:0 auto;padding:24px 20px;">
    <div style="background:#e3f2fd;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(13,71,161,0.08);border-left:5px solid #1976d2;">
      <div style="padding:22px 24px 18px;">
        <p style="margin:0 0 10px;font-size:15px;color:#0d47a1;">Hello ${escapeHtml(userName || '')},</p>
        <p style="margin:0 0 18px;font-size:14px;color:#37474f;">Here is your Task View summary:</p>
        <table style="width:100%;border-collapse:collapse;font-size:14px;background:#fff;border-radius:8px;overflow:hidden;">
          <thead>
            <tr style="background:#bbdefb;color:#0d47a1;">
              <th style="text-align:left;padding:12px 14px;font-weight:600;">Task</th>
              <th style="text-align:left;padding:12px 14px;font-weight:600;">Due</th>
              <th style="text-align:left;padding:12px 14px;font-weight:600;">Status</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
        <p style="margin:16px 0 0;font-size:12px;color:#546e7a;">Generated ${escapeHtml(istLine)} (IST)</p>
      </div>
    </div>
    ${emailBrandFooterHtml()}
  </div>
</body></html>`;
}

async function sendTaskViewSummaryEmail(toEmail, userName, tasks) {
  if (!isEmailEnabled()) {
    throw new Error('Email not configured');
  }
  if (!Array.isArray(tasks) || tasks.length === 0) {
    throw new Error('No tasks');
  }
  const rows = buildTaskViewSummaryTableRows(tasks);
  const istLine = formatIstGeneratedLine();
  const defaultSubject = `Task View summary (${tasks.length} task${tasks.length === 1 ? '' : 's'})`;
  const defaultBody = buildTaskViewSummaryEmailHtml(userName, rows, istLine);
  const tpl = await getEmailTemplatePair('task_view_summary');
  const plain = {
    userName: userName || '',
    taskCount: String(tasks.length),
    generatedLine: istLine,
  };
  const htmlVars = {
    ...plain,
    userName: escapeHtml(plain.userName),
    generatedLine: escapeHtml(istLine),
    taskRows: rows,
    defaultBody,
  };
  const subject = tpl.subject ? interpolateEmailTemplate(tpl.subject, plain) : defaultSubject;
  const html = tpl.bodyHtml ? interpolateEmailTemplate(tpl.bodyHtml, htmlVars) : defaultBody;
  await sendMail(toEmail, subject, html);
}

function buildNeedImprovementFinalizedHtml(assigneeName, taskTitle, adminComment, adminName, rowsHtml, istLine) {
  const st = taskViewEmailStatusStyle('Needs Improvement');
  const row =
    rowsHtml ||
    `<tr>
<td style="padding:10px 12px;border-bottom:1px solid #bbdefb;vertical-align:middle;">${escapeHtml(taskTitle || '(untitled)')}</td>
<td style="padding:10px 12px;border-bottom:1px solid #bbdefb;vertical-align:middle;color:#1565c0;">—</td>
<td style="padding:10px 12px;border-bottom:1px solid #bbdefb;vertical-align:middle;font-weight:600;background:${st.bg};color:${st.color};">Needs Improvement</td>
</tr>`;
  return `
<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#e8f4fc;">
  <div style="font-family:Segoe UI,Roboto,Helvetica,sans-serif;max-width:720px;margin:0 auto;padding:24px 20px;">
    <div style="background:#e3f2fd;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(13,71,161,0.08);border-left:5px solid #1976d2;">
      <div style="padding:22px 24px 18px;">
        <p style="margin:0 0 10px;font-size:15px;color:#0d47a1;">Hello ${escapeHtml(assigneeName || '')},</p>
        <p style="margin:0 0 14px;font-size:14px;color:#37474f;">Your task completion was reviewed and marked <strong>Needs improvement</strong>. Admin comment:</p>
        <p style="margin:0 0 18px;font-size:15px;font-weight:700;color:#5d4037;white-space:pre-wrap;">${escapeHtml(adminComment || '')}</p>
        <p style="margin:0 0 8px;font-size:13px;color:#546e7a;">Reviewer: ${escapeHtml(adminName || 'Admin')}</p>
        <table style="width:100%;border-collapse:collapse;font-size:14px;background:#fff;border-radius:8px;overflow:hidden;">
          <thead>
            <tr style="background:#bbdefb;color:#0d47a1;">
              <th style="text-align:left;padding:12px 14px;font-weight:600;">Task</th>
              <th style="text-align:left;padding:12px 14px;font-weight:600;">Due</th>
              <th style="text-align:left;padding:12px 14px;font-weight:600;">Status</th>
            </tr>
          </thead>
          <tbody>${row}</tbody>
        </table>
        <p style="margin:16px 0 0;font-size:12px;color:#546e7a;">Generated ${escapeHtml(istLine)} (IST)</p>
      </div>
    </div>
    ${emailBrandFooterHtml()}
  </div>
</body></html>`;
}

async function sendNeedImprovementFinalizedEmail(toEmail, assigneeName, taskTitle, adminComment, adminName) {
  if (!isEmailEnabled()) return false;
  const istLine = formatIstGeneratedLine();
  const rows = buildTaskViewSummaryTableRows([
    { title: taskTitle || '(untitled)', due: '—', status: 'Needs Improvement' },
  ]);
  const defaultSubject = `Task reviewed — needs improvement: ${taskTitle || '(untitled)'}`;
  const defaultBody = buildNeedImprovementFinalizedHtml(
    assigneeName,
    taskTitle,
    adminComment,
    adminName,
    rows,
    istLine
  );
  const tpl = await getEmailTemplatePair('need_improvement_finalized');
  const plain = {
    assigneeName: assigneeName || '',
    taskTitle: taskTitle || '(untitled)',
    adminComment: adminComment || '',
    adminName: adminName || 'Admin',
  };
  const htmlVars = {
    ...plain,
    assigneeName: escapeHtml(plain.assigneeName),
    taskTitle: escapeHtml(plain.taskTitle),
    adminComment: escapeHtml(plain.adminComment),
    adminName: escapeHtml(plain.adminName),
    generatedLine: escapeHtml(istLine),
    taskRows: rows,
    defaultBody,
  };
  const subject = tpl.subject ? interpolateEmailTemplate(tpl.subject, plain) : defaultSubject;
  const html = tpl.bodyHtml ? interpolateEmailTemplate(tpl.bodyHtml, htmlVars) : defaultBody;
  try {
    await sendMail(toEmail, subject, html);
    return true;
  } catch (err) {
    console.error(`Need improvement finalized email failed for ${toEmail}:`, err.message);
    return false;
  }
}

module.exports = {
  isEmailEnabled,
  sendTaskReminderEmail,
  sendTaskAssignmentEmail,
  sendTestEmail,
  sendTaskRejectedEmail,
  sendAccountCreatedEmail,
  sendPasswordResetCodeEmail,
  sendTaskViewSummaryEmail,
  sendNeedImprovementFinalizedEmail,
};
