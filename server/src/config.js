require('dotenv').config();

const PORT = parseInt(process.env.PORT || '4000', 10);
const JWT_SECRET = process.env.JWT_SECRET || 'dev-only-change-me';
const JWT_EXPIRES = process.env.JWT_EXPIRES || '7d';
const MASTER_EMAIL = (process.env.MASTER_EMAIL || 'mansuri.amin1@gmail.com').toLowerCase().trim();
const MASTER_PASSWORD = process.env.MASTER_PASSWORD || '';
const MONGODB_URI = process.env.MONGODB_URI || '';

function parseOrigins() {
  const raw = process.env.FRONTEND_ORIGIN || '';
  if (!raw.trim()) {
    return ['http://localhost:5173', 'http://localhost:8080', 'http://127.0.0.1:5500', 'http://localhost:5500'];
  }
  return raw.split(',').map((o) => o.trim()).filter(Boolean);
}

const SMTP_EMAIL = process.env.SMTP_EMAIL || '';
const SMTP_PASSWORD = process.env.SMTP_PASSWORD || '';
const SMTP_HOST = process.env.SMTP_HOST || 'smtp.gmail.com';
const SMTP_PORT = parseInt(process.env.SMTP_PORT || '587', 10);
const REMINDER_CRON = process.env.REMINDER_CRON || '0 8 * * *';
const SMTP_CONFIGURED = !!(SMTP_EMAIL && SMTP_PASSWORD);

module.exports = {
  PORT,
  JWT_SECRET,
  JWT_EXPIRES,
  MASTER_EMAIL,
  MASTER_PASSWORD,
  MONGODB_URI,
  parseOrigins,
  isProduction: process.env.NODE_ENV === 'production',
  SMTP_EMAIL,
  SMTP_PASSWORD,
  SMTP_HOST,
  SMTP_PORT,
  REMINDER_CRON,
  SMTP_CONFIGURED,
};
