const mongoose = require('mongoose');

const workspaceSchema = new mongoose.Schema(
  {
    tenantRootUserId: { type: Number, sparse: true, unique: true, index: true },
    data: { type: mongoose.Schema.Types.Mixed, required: true },
  },
  { timestamps: true }
);

module.exports = mongoose.model('Workspace', workspaceSchema);
