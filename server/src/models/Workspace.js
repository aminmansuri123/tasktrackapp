const mongoose = require('mongoose');

const workspaceSchema = new mongoose.Schema(
  {
    /** @deprecated Legacy single-tenant key */
    name: { type: String },
    /** Primary admin userId for this isolated workspace (one document per org) */
    tenantRootUserId: { type: Number, sparse: true, unique: true, index: true },
    data: { type: mongoose.Schema.Types.Mixed, required: true },
  },
  { timestamps: true }
);

module.exports = mongoose.model('Workspace', workspaceSchema);
