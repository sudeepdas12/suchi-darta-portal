import B2 from 'backblaze-b2';

const ADMIN_TOKEN = process.env.ADMIN_TOKEN || '';

function isAuthorized(token) {
  if (!ADMIN_TOKEN) return true;
  return token === ADMIN_TOKEN;
}

export default async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  const token = req.query.token;
  if (!isAuthorized(token)) return res.status(403).json({ error: 'Unauthorized' });

  const { B2_KEY_ID, B2_APP_KEY } = process.env;
  if (!B2_KEY_ID || !B2_APP_KEY) return res.status(500).json({ error: 'Missing Backblaze B2 environment variables' });

  const fileId = req.query.fileId;
  if (!fileId) return res.status(400).json({ error: 'Missing fileId' });

  const b2 = new B2({ applicationKeyId: B2_KEY_ID, applicationKey: B2_APP_KEY });
  try {
    await b2.authorize();
    // Attempt to get file info for filename
    let filename = req.query.name || '';
    try {
      const info = await b2.getFileInfo({ fileId });
      if (!filename && info && info.data && info.data.fileName) {
        filename = info.data.fileName;
      }
    } catch (e) {
      // ignore
    }

    const downloadResp = await b2.downloadFileById({ fileId, responseType: 'arraybuffer' });
    const buffer = Buffer.from(downloadResp.data);

    res.setHeader('Content-Type', 'application/octet-stream');
    const safeName = filename ? filename.split('/').pop() : `attachment-${fileId}`;
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}"`);
    return res.status(200).send(buffer);
  } catch (err) {
    return res.status(500).json({ error: err.message || 'Unable to download attachment' });
  }
}
