import B2 from 'backblaze-b2';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { B2_KEY_ID, B2_APP_KEY, B2_BUCKET_ID } = process.env;
  if (!B2_KEY_ID || !B2_APP_KEY || !B2_BUCKET_ID) {
    return res.status(500).json({
      error: 'Missing Backblaze B2 environment variables: B2_KEY_ID, B2_APP_KEY, B2_BUCKET_ID'
    });
  }

  const payload = req.body;
  if (!payload || typeof payload !== 'object') {
    return res.status(400).json({ error: 'Missing registration payload' });
  }

  const b2 = new B2({
    applicationKeyId: B2_KEY_ID,
    applicationKey: B2_APP_KEY
  });

  try {
    await b2.authorize();

    const response = await b2.getUploadUrl({ bucketId: B2_BUCKET_ID });
    const fileName = `registrations/${Date.now()}-${Math.random().toString(36).slice(2, 10)}.json`;
    const fileContent = JSON.stringify({ ...payload, submittedAt: new Date().toISOString() });

    const uploadResponse = await fetch(response.data.uploadUrl, {
      method: 'POST',
      headers: {
        Authorization: response.data.authorizationToken,
        'X-Bz-File-Name': encodeURIComponent(fileName),
        'Content-Type': 'application/json',
        'X-Bz-Content-Sha1': 'do_not_verify'
      },
      body: fileContent
    });

    if (!uploadResponse.ok) {
      const errorText = await uploadResponse.text().catch(() => 'Unknown upload error');
      throw new Error(`Backblaze upload failed: ${uploadResponse.status} ${errorText}`);
    }

    const uploadResult = await uploadResponse.json();
    return res.status(200).json({ success: true, fileName, fileId: uploadResult.fileId });
  } catch (error) {
    return res.status(500).json({ error: error.message || 'Unable to save registration' });
  }
}
