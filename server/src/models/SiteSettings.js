const mongoose = require('mongoose');

const siteSettingsSchema = new mongoose.Schema(
  {
    registrationMode: {
      type: String,
      enum: ['open', 'email_list', 'domain_list'],
      default: 'open',
    },
    /** Lowercased emails allowed to self-register when registrationMode is email_list */
    allowedEmails: { type: [String], default: [] },
    /** Lowercased hostnames (no @) allowed when registrationMode is domain_list */
    allowedDomains: { type: [String], default: [] },
  },
  { timestamps: true }
);

module.exports = mongoose.model('SiteSettings', siteSettingsSchema);
