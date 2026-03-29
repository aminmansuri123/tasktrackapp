const dns = require('dns');
const nodemailer = require('nodemailer');
const { SMTP_EMAIL, SMTP_PASSWORD, SMTP_HOST, SMTP_PORT, SMTP_CONFIGURED } = require('../config');

// Force IPv4 DNS resolution — many cloud providers (Render, Railway, etc.)
// don't support outbound IPv6, causing ENETUNREACH errors with Gmail SMTP.
dns.setDefaultResultOrder('ipv4first');

function isEmailEnabled() {
  return SMTP_CONFIGURED;
}

let _transporter = null;

function getTransporter() {
  if (!SMTP_CONFIGURED) return null;
  if (_transporter) return _transporter;
  _transporter = nodemailer.createTransport({
    host: SMTP_HOST,
    port: SMTP_PORT,
    secure: SMTP_PORT === 465,
    auth: { user: SMTP_EMAIL, pass: SMTP_PASSWORD },
  });
  return _transporter;
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
    </div>
  </div>
</body>
</html>`;
}

async function sendTaskReminderEmail(toEmail, userName, overdueTasks, upcomingTasks) {
  if (!isEmailEnabled()) return false;
  const transporter = getTransporter();
  if (!transporter) return false;

  const total = overdueTasks.length + upcomingTasks.length;
  const subject = total === 1
    ? 'Task Reminder: 1 task needs attention'
    : `Task Reminder: ${total} tasks need attention`;

  try {
    await transporter.sendMail({
      from: `"Task Tracker" <${SMTP_EMAIL}>`,
      to: toEmail,
      subject,
      html: buildReminderHtml(userName, overdueTasks, upcomingTasks),
    });
    return true;
  } catch (err) {
    console.error(`Email send failed for ${toEmail}:`, err.message);
    return false;
  }
}

function buildAssignmentHtml(assigneeName, taskTitle, dueDate, assignerName, isSelf) {
  const dueLine = dueDate ? `<strong>Due:</strong> ${formatDate(dueDate)}` : '<strong>Due:</strong> Not set';
  const byLine = isSelf
    ? 'You created this task for yourself.'
    : `Assigned by <strong>${assignerName || 'an admin'}</strong>.`;
  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;color:#333;margin:0;padding:20px;background:#f5f5f5;">
  <div style="max-width:640px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">
    <div style="background:#16a34a;color:#fff;padding:20px 24px;">
      <h1 style="margin:0;font-size:20px;">New Task Assigned</h1>
      <p style="margin:6px 0 0;font-size:14px;opacity:0.9;">Hello ${assigneeName}, a task has been ${isSelf ? 'created by you' : 'assigned to you'}.</p>
    </div>
    <div style="padding:20px 24px;">
      <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:16px;">
        <tr><td style="padding:8px 0;font-weight:600;width:100px;">Task:</td><td style="padding:8px 0;">${taskTitle || '(untitled)'}</td></tr>
        <tr><td style="padding:8px 0;font-weight:600;">Due Date:</td><td style="padding:8px 0;">${dueDate ? formatDate(dueDate) : 'Not set'}</td></tr>
        <tr><td style="padding:8px 0;font-weight:600;">Assigned By:</td><td style="padding:8px 0;">${isSelf ? 'Self' : (assignerName || 'Admin')}</td></tr>
      </table>
      <p style="font-size:13px;color:#666;">${byLine}</p>
      <p style="font-size:13px;color:#999;margin-top:16px;">This is an automated notification from your Task Management System.</p>
    </div>
  </div>
</body>
</html>`;
}

async function sendTaskAssignmentEmail(toEmail, assigneeName, taskTitle, dueDate, assignerName, isSelf) {
  if (!isEmailEnabled()) return false;
  const transporter = getTransporter();
  if (!transporter) return false;

  const subject = isSelf
    ? `New Task: ${taskTitle || '(untitled)'}`
    : `Task Assigned: ${taskTitle || '(untitled)'}`;

  try {
    await transporter.sendMail({
      from: `"Task Tracker" <${SMTP_EMAIL}>`,
      to: toEmail,
      subject,
      html: buildAssignmentHtml(assigneeName, taskTitle, dueDate, assignerName, isSelf),
    });
    return true;
  } catch (err) {
    console.error(`Assignment email failed for ${toEmail}:`, err.message);
    return false;
  }
}

async function sendTestEmail(toEmail, userName) {
  if (!isEmailEnabled()) throw new Error('SMTP not configured');
  const transporter = getTransporter();
  if (!transporter) throw new Error('Could not create email transporter');
  const demoTask = {
    title: 'Sample Task — Test Reminder',
    due_date: new Date().toISOString().split('T')[0],
    status: 'in_progress',
  };
  await transporter.sendMail({
    from: `"Task Tracker" <${SMTP_EMAIL}>`,
    to: toEmail,
    subject: 'Test Email — Task Tracker Reminder',
    html: buildReminderHtml(userName, [], [demoTask]),
  });
  return true;
}

module.exports = { isEmailEnabled, sendTaskReminderEmail, sendTaskAssignmentEmail, sendTestEmail };
