const express = require('express');
const multer = require('multer');
const { authMiddleware } = require('../middleware/auth');
const {
  uploadAttachmentBuffer,
  openDownloadStream,
  deleteAttachmentFile,
} = require('../services/gridfsAttachments');

const router = express.Router();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

router.post('/', authMiddleware, upload.single('file'), async (req, res) => {
  try {
    const locationId = req.body.locationId;
    const attachmentId = req.body.attachmentId;
    if (locationId === undefined || attachmentId === undefined || !req.file) {
      return res.status(400).json({ error: 'locationId, attachmentId, and file required' });
    }
    await uploadAttachmentBuffer(locationId, attachmentId, req.file.buffer, req.file.mimetype);
    return res.status(201).json({ ok: true, storedRemotely: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Upload failed' });
  }
});

router.get('/:locationId/:attachmentId', authMiddleware, async (req, res) => {
  try {
    const { locationId, attachmentId } = req.params;
    const result = await openDownloadStream(locationId, attachmentId);
    if (!result) {
      return res.status(404).json({ error: 'Not found' });
    }
    res.setHeader('Content-Type', result.contentType);
    result.stream.on('error', () => {
      if (!res.headersSent) res.status(500).end();
    });
    result.stream.pipe(res);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Download failed' });
  }
});

router.delete('/:locationId/:attachmentId', authMiddleware, async (req, res) => {
  try {
    const { locationId, attachmentId } = req.params;
    await deleteAttachmentFile(locationId, attachmentId);
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Delete failed' });
  }
});

module.exports = router;
