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

  const b2 = new B2({
    applicationKeyId: B2_KEY_ID,
    applicationKey: B2_APP_KEY
  });

  try {
    await b2.authorize();

    let files = [];
    let nextFileName;

    do {
      const response = await b2.listFileNames({
        bucketId: B2_BUCKET_ID,
        startFileName: nextFileName,
        maxFileCount: 100
      });

      const batch = response.data.files || [];
      files.push(...batch);
      nextFileName = response.data.nextFileName;
    } while (nextFileName);

    const summaryFiles = files
      .filter(f => f.fileName.toLowerCase().endsWith('.xlsx') && f.fileName.includes('registration-'))
      .map(f => ({ fileName: f.fileName, fileId: f.fileId, uploadTimestamp: f.uploadTimestamp }));

    return res.status(200).json({ summaries: summaryFiles });
  } catch (error) {
    return res.status(500).json({ error: error.message || 'Unable to list summary files' });
  }
}
