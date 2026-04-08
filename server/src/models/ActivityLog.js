const mongoose = require('mongoose');

const activityLogSchema = new mongoose.Schema(
  {
    actorUserId: { type: Number, required: true, index: true },
    tenantRootUserId: { type: Number, default: null, index: true },
    action: { type: String, required: true, maxlength: 64 },
    entityType: { type: String, maxlength: 32, default: '' },
    entityId: { type: String, maxlength: 64, default: '' },
    summary: { type: String, maxlength: 500, default: '' },
  },
  { timestamps: { createdAt: 'createdAt', updatedAt: false } }
);

activityLogSchema.index({ actorUserId: 1, createdAt: -1 });

module.exports = mongoose.model('ActivityLog', activityLogSchema);
