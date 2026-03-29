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

function isEmailEnabled() {
  return EMAIL_CONFIGURED;
}

function resendFromHeader() {
  if (EMAIL_FROM) return EMAIL_FROM;
  return 'Task Tracker <onboarding@resend.dev>';
}

async function sendMailResend(toEmail, subject, html) {
  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${RESEND_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from: resendFromHeader(),
      to: [toEmail],
      subject,
      html,
    }),
  });
  const bodyText = await res.text();
  if (!res.ok) {
    throw new Error(`Resend API ${res.status}: ${bodyText.slice(0, 500)}`);
  }
}

async function sendMailSmtp(toEmail, subject, html) {
  const transporter = await getTransporter();
  if (!transporter) throw new Error('SMTP transporter unavailable');
  await transporter.sendMail({
    from: `"Task Tracker" <${SMTP_EMAIL}>`,
    to: toEmail,
    subject,
    html,
  });
}

let _loggedTransport = false;
async function sendMail(toEmail, subject, html) {
  if (RESEND_API_KEY) {
    if (!_loggedTransport) {
      console.log('[email] Sending via Resend API (HTTPS)');
      _loggedTransport = true;
    }
    await sendMailResend(toEmail, subject, html);
    return;
  }
  if (!_loggedTransport) {
    console.log('[email] Sending via SMTP');
    _loggedTransport = true;
  }
  await sendMailSmtp(toEmail, subject, html);
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
  const subject = total === 1
    ? 'Task Reminder: 1 task needs attention'
    : `Task Reminder: ${total} tasks need attention`;

  try {
    await sendMail(toEmail, subject, buildReminderHtml(userName, overdueTasks, upcomingTasks));
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
  let subject;
  if (eventKind === 'reassigned' && !isSelf) {
    subject = `Task reassigned to you: ${title}`;
  } else if (isSelf) {
    subject = `New task created: ${title}`;
  } else {
    subject = `New task assigned to you: ${title}`;
  }

  try {
    await sendMail(toEmail, subject, buildAssignmentHtml(assigneeName, taskTitle, dueDate, assignerName, isSelf));
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
  const subject = `Task needs revision: ${taskTitle || '(untitled)'}`;
  try {
    await sendMail(toEmail, subject, buildTaskRejectedHtml(assigneeName, taskTitle, comment, adminName));
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
  try {
    await sendMail(toEmail, 'Your Task Tracker account is ready', buildAccountCreatedHtml(userName, contextHtml));
    return true;
  } catch (err) {
    console.error(`Account created email failed for ${toEmail}:`, err.message);
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

module.exports = {
  isEmailEnabled,
  sendTaskReminderEmail,
  sendTaskAssignmentEmail,
  sendTestEmail,
  sendTaskRejectedEmail,
  sendAccountCreatedEmail,
};
