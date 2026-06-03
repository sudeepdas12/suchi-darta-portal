// api/get-upload-url.js
import B2 from 'backblaze-b2';

export default async function handler(req, res) {
  // Only allow secure POST traffic
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { B2_KEY_ID, B2_APP_KEY, B2_BUCKET_ID } = process.env;
  if (!B2_KEY_ID || !B2_APP_KEY || !B2_BUCKET_ID) {
    return res.status(500).json({
      error: 'Missing Backblaze B2 environment variables: B2_KEY_ID, B2_APP_KEY, B2_BUCKET_ID'
    });
  }

  const b2 = new B2({
    applicationKeyId: B2_KEY_ID,
    applicationKey: B2_APP_KEY
  });

  try {
    // 1. Handshake with Backblaze API
    await b2.authorize();

    // 2. Request a one-time unique token to allow file stream entry
    const response = await b2.getUploadUrl({
      bucketId: process.env.B2_BUCKET_ID
    });

    // 3. Return the authorization strings to the waiting browser form
    return res.status(200).json({
      uploadUrl: response.data.uploadUrl,
      uploadAuthToken: response.data.authorizationToken
    });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}
