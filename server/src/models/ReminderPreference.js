const mongoose = require('mongoose');

const reminderPreferenceSchema = new mongoose.Schema(
  {
    userId: { type: Number, required: true, unique: true, index: true },
    beforeDueDate: { type: Boolean, default: true },
    afterDueDate: { type: Boolean, default: true },
    setByAdmin: { type: Boolean, default: false },
  },
  { timestamps: true }
);

module.exports = mongoose.model('ReminderPreference', reminderPreferenceSchema);
