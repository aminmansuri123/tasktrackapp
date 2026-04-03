/**
 * Last-login display for admin / workspace user lists — India Standard Time (UTC+05:30).
 * Format: DD-MM-YYYY HH:MM:SS (24h).
 */
const APP_TIMEZONE = 'Asia/Kolkata';

function formatLastLoginAtDisplay(d) {
  if (!d) return '';
  try {
    const date = d instanceof Date ? d : new Date(d);
    if (Number.isNaN(date.getTime())) return '';
    const fmt = new Intl.DateTimeFormat('en-GB', {
      timeZone: APP_TIMEZONE,
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false,
    });
    const parts = fmt.formatToParts(date);
    const get = (t) => parts.find((p) => p.type === t)?.value ?? '';
    return `${get('day')}-${get('month')}-${get('year')} ${get('hour')}:${get('minute')}:${get('second')}`;
  } catch {
    return '';
  }
}

module.exports = { formatLastLoginAtDisplay, APP_TIMEZONE };
