const mongoose = require('mongoose');

const userSchema = new mongoose.Schema(
  {
    userId: { type: Number, required: true, unique: true, index: true },
    email: { type: String, required: true, unique: true, lowercase: true, trim: true },
    passwordHash: { type: String, required: true },
    name: { type: String, required: true, trim: true },
    role: { type: String, enum: ['admin', 'user'], default: 'user' },
    isActive: { type: Boolean, default: true },
    isMaster: { type: Boolean, default: false },
    /** Org root admin userId; null for master account only. All tenant members share the same value. */
    tenantRootUserId: { type: Number, default: null, index: true },
    /** Feature flags toggled by master: 'locations', 'codeSnippets' */
    enabledFeatures: { type: [String], default: [] },
    /** Additional tenant roots this user is shared with (visible + assignable in those orgs). */
    sharedWithTenants: { type: [Number], default: [] },
    /** Ephemeral state for email-based password reset (4-digit code). */
    passwordResetCodeHash: { type: String, default: null },
    passwordResetExpiresAt: { type: Date, default: null },
    passwordResetAttempts: { type: Number, default: 0 },
    passwordResetLastSentAt: { type: Date, default: null },
  },
  { timestamps: true }
);

module.exports = mongoose.model('User', userSchema);
