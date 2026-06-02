// api/get-upload-url.js
import B2 from 'backblaze-b2';

// Initialize connection using system variables hidden on Vercel's servers
const b2 = new B2({
  applicationKeyId: process.env.B2_KEY_ID, 
  applicationKey: process.env.B2_APP_KEY
});

export default async function handler(req, res) {
  // Only allow secure POST traffic
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

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
