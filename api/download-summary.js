import B2 from 'backblaze-b2';

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { B2_KEY_ID, B2_APP_KEY, B2_BUCKET_ID } = process.env;
  if (!B2_KEY_ID || !B2_APP_KEY || !B2_BUCKET_ID) {
    return res.status(500).json({ error: 'Missing Backblaze B2 environment variables' });
  }

  const fileId = req.query.fileId;
  const downloadName = req.query.name || 'registration-summary.xlsx';
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
    res.setHeader('Content-Disposition', `attachment; filename="${downloadName.replace(/[^a-zA-Z0-9._-]/g, '_')}"`);
    return res.status(200).send(fileBuffer);
  } catch (error) {
    return res.status(500).json({ error: error.message || 'Unable to download summary file' });
  }
}
