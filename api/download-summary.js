import B2 from 'backblaze-b2';

const ADMIN_TOKEN = process.env.ADMIN_TOKEN || '';

function isAuthorized(token) {
  if (!ADMIN_TOKEN) {
    return true;
  }
  return token === ADMIN_TOKEN;
}

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const token = req.query.token;
  if (!isAuthorized(token)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }

  const { B2_KEY_ID, B2_APP_KEY, B2_BUCKET_ID } = process.env;
  if (!B2_KEY_ID || !B2_APP_KEY || !B2_BUCKET_ID) {
    return res.status(500).json({ error: 'Missing Backblaze B2 environment variables' });
  }

  const fileId = req.query.fileId;
  if (!fileId) {
    return res.status(400).json({ error: 'Missing fileId query parameter' });
  }

  const b2 = new B2({
    applicationKeyId: B2_KEY_ID,
    applicationKey: B2_APP_KEY
  });

  try {
    await b2.authorize();
    const downloadResponse = await b2.downloadFileById({
      fileId,
      responseType: 'arraybuffer'
    });

    const fileBuffer = Buffer.from(downloadResponse.data);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="registration-summary.xlsx"');
    return res.status(200).send(fileBuffer);
  } catch (error) {
    return res.status(500).json({ error: error.message || 'Unable to download summary file' });
  }
}
