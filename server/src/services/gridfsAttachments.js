const mongoose = require('mongoose');

const BUCKET_NAME = 'locationAttachments';

function getBucket() {
  const db = mongoose.connection.db;
  if (!db) return null;
  return new mongoose.mongo.GridFSBucket(db, { bucketName: BUCKET_NAME });
}

function attachmentFilename(locationId, attachmentId) {
  return `${locationId}_${attachmentId}`;
}

async function deleteAttachmentFile(locationId, attachmentId) {
  const bucket = getBucket();
  if (!bucket) return;
  const filename = attachmentFilename(locationId, attachmentId);
  const cursor = bucket.find({ filename });
  // eslint-disable-next-line no-restricted-syntax
  for await (const doc of cursor) {
    // eslint-disable-next-line no-await-in-loop
    await bucket.delete(doc._id);
  }
}

async function uploadAttachmentBuffer(locationId, attachmentId, buffer, contentType) {
  const bucket = getBucket();
  if (!bucket) throw new Error('GridFS not ready');
  await deleteAttachmentFile(locationId, attachmentId);
  const filename = attachmentFilename(locationId, attachmentId);
  return new Promise((resolve, reject) => {
    const upload = bucket.openUploadStream(filename, {
      contentType: contentType || 'application/octet-stream',
      metadata: { locationId: String(locationId), attachmentId: String(attachmentId) },
    });
    upload.on('error', reject);
    upload.on('finish', () => resolve(upload.id));
    upload.end(buffer);
  });
}

async function openDownloadStream(locationId, attachmentId) {
  const bucket = getBucket();
  if (!bucket) return null;
  const filename = attachmentFilename(locationId, attachmentId);
  const files = await bucket.find({ filename }).limit(1).toArray();
  if (!files.length) return null;
  return {
    stream: bucket.openDownloadStream(files[0]._id),
    contentType: files[0].contentType || 'application/octet-stream',
    filename: files[0].filename,
  };
}

module.exports = {
  uploadAttachmentBuffer,
  openDownloadStream,
  deleteAttachmentFile,
  attachmentFilename,
};
