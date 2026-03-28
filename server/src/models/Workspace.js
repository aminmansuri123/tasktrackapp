const mongoose = require('mongoose');

const workspaceSchema = new mongoose.Schema(
  {
    name: { type: String, default: 'default', unique: true },
    data: { type: mongoose.Schema.Types.Mixed, required: true },
  },
  { timestamps: true }
);

module.exports = mongoose.model('Workspace', workspaceSchema);
