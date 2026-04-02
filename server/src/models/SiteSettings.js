const mongoose = require('mongoose');

const siteSettingsSchema = new mongoose.Schema(
  {
    registrationMode: {
      type: String,
      enum: ['open', 'restricted', 'email_list', 'domain_list'],
      default: 'open',
    },
    /** Lowercased emails allowed to self-register */
    allowedEmails: { type: [String], default: [] },
    /** Lowercased hostnames (no @) allowed */
    allowedDomains: { type: [String], default: [] },
    /** 0 = disabled. Master-only; enforced in auth middleware with User.lastActivityAt. */
    sessionIdleTimeoutMinutes: { type: Number, default: 0 },
  },
  { timestamps: true }
);

module.exports = mongoose.model('SiteSettings', siteSettingsSchema);
