const cron = require('node-cron');
const { REMINDER_CRON } = require('../config');
const { isEmailEnabled, sendTaskReminderEmail } = require('./emailService');
const User = require('../models/User');
const Workspace = require('../models/Workspace');
const ReminderPreference = require('../models/ReminderPreference');

function todayStr() {
  const d = new Date();
  return d.toISOString().split('T')[0];
}

function tomorrowStr() {
  const d = new Date();
  d.setDate(d.getDate() + 1);
  return d.toISOString().split('T')[0];
}

async function runReminders() {
  if (!isEmailEnabled()) return;
  console.log('[reminder-cron] Running daily task reminders…');

  try {
    const today = todayStr();
    const tomorrow = tomorrowStr();
    const allUsers = await User.find({ isActive: true, isMaster: { $ne: true } }).lean();
    const allPrefs = await ReminderPreference.find({}).lean();
    const prefsMap = new Map(allPrefs.map((p) => [p.userId, p]));

    for (const user of allUsers) {
      const pref = prefsMap.get(user.userId) || { beforeDueDate: true, afterDueDate: true };
      if (!pref.beforeDueDate && !pref.afterDueDate) continue;

      const tenantRoot = user.tenantRootUserId ?? user.userId;
      const ws = await Workspace.findOne({ tenantRootUserId: tenantRoot }).lean();
      if (!ws || !ws.data || !Array.isArray(ws.data.tasks)) continue;

      const userTasks = ws.data.tasks.filter(
        (t) => Number(t.assigned_to) === user.userId
      );

      const overdue = [];
      const upcoming = [];

      for (const t of userTasks) {
        if (t.status === 'completed') continue;
        const due = t.due_date || t.next_due_date;
        if (!due) continue;
        if (pref.afterDueDate && due < today) overdue.push(t);
        if (pref.beforeDueDate && due === tomorrow) upcoming.push(t);
      }

      if (overdue.length === 0 && upcoming.length === 0) continue;

      const ok = await sendTaskReminderEmail(user.email, user.name || user.email, overdue, upcoming);
      if (ok) {
        console.log(`[reminder-cron] Sent to ${user.email}: ${overdue.length} overdue, ${upcoming.length} upcoming`);
      }
    }

    console.log('[reminder-cron] Daily reminders complete.');
  } catch (err) {
    console.error('[reminder-cron] Error:', err);
  }
}

function startReminderCron() {
  if (!isEmailEnabled()) {
    console.log('[reminder-cron] SMTP not configured — skipping cron setup.');
    return;
  }
  if (!cron.validate(REMINDER_CRON)) {
    console.error(`[reminder-cron] Invalid cron expression: ${REMINDER_CRON}`);
    return;
  }
  cron.schedule(REMINDER_CRON, runReminders, { timezone: 'Asia/Kolkata' });
  console.log(`[reminder-cron] Scheduled with cron "${REMINDER_CRON}" (Asia/Kolkata)`);
}

module.exports = { startReminderCron, runReminders };
