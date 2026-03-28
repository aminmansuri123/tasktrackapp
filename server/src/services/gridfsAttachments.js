const mongoose = require('mongoose');

const BUCKET_NAME = 'locationAttachments';

function getBucket() {
  const db = mongoose.connection.db;
  if (!db) return null;
  return new mongoose.mongo.GridFSBucket(db, { bucketName: BUCKET_NAME });
}

function attachmentFilenameLegacy(locationId, attachmentId) {
  return `${locationId}_${attachmentId}`;
}

/** Tenant-scoped filename (org isolation). */
function attachmentFilename(tenantKey, locationId, attachmentId) {
  const t = tenantKey == null || tenantKey === 'legacy' ? 'legacy' : String(tenantKey);
  return `${t}_${locationId}_${attachmentId}`;
}

async function deleteAttachmentFile(tenantKey, locationId, attachmentId) {
  const bucket = getBucket();
  if (!bucket) return;
  const names = [attachmentFilename(tenantKey, locationId, attachmentId), attachmentFilenameLegacy(locationId, attachmentId)];
  const uniq = [...new Set(names)];
  // eslint-disable-next-line no-restricted-syntax
  for (const filename of uniq) {
    const cursor = bucket.find({ filename });
    // eslint-disable-next-line no-restricted-syntax
    for await (const doc of cursor) {
      // eslint-disable-next-line no-await-in-loop
      await bucket.delete(doc._id);
    }
  }
}

async function uploadAttachmentBuffer(tenantKey, locationId, attachmentId, buffer, contentType) {
  const bucket = getBucket();
  if (!bucket) throw new Error('GridFS not ready');
  await deleteAttachmentFile(tenantKey, locationId, attachmentId);
  const filename = attachmentFilename(tenantKey, locationId, attachmentId);
  return new Promise((resolve, reject) => {
    const upload = bucket.openUploadStream(filename, {
      contentType: contentType || 'application/octet-stream',
      metadata: {
        tenantKey: String(tenantKey),
        locationId: String(locationId),
        attachmentId: String(attachmentId),
      },
    });
    upload.on('error', reject);
    upload.on('finish', () => resolve(upload.id));
    upload.end(buffer);
  });
}

async function openDownloadStream(tenantKey, locationId, attachmentId) {
  const bucket = getBucket();
  if (!bucket) return null;
  const candidates = [
    attachmentFilename(tenantKey, locationId, attachmentId),
    attachmentFilenameLegacy(locationId, attachmentId),
  ];
  const uniq = [...new Set(candidates)];
  // eslint-disable-next-line no-restricted-syntax
  for (const filename of uniq) {
    // eslint-disable-next-line no-await-in-loop
    const files = await bucket.find({ filename }).limit(1).toArray();
    if (files.length) {
      return {
        stream: bucket.openDownloadStream(files[0]._id),
        contentType: files[0].contentType || 'application/octet-stream',
        filename: files[0].filename,
      };
    }
  }
  return null;
}

module.exports = {
  uploadAttachmentBuffer,
  openDownloadStream,
  deleteAttachmentFile,
  attachmentFilename,
  attachmentFilenameLegacy,
};
