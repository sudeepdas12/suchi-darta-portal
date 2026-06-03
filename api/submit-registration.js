import { promises as fs } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const storageFile = path.join(__dirname, '..', 'registrations.json');

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const payload = req.body;
    if (!payload || typeof payload !== 'object') {
      throw new Error('Missing registration payload');
    }

    const raw = await fs.readFile(storageFile, 'utf8').catch(() => '[]');
    const registrations = Array.isArray(JSON.parse(raw)) ? JSON.parse(raw) : [];

    registrations.push({
      ...payload,
      submittedAt: new Date().toISOString()
    });

    await fs.writeFile(storageFile, JSON.stringify(registrations, null, 2), 'utf8');

    return res.status(200).json({ success: true });
  } catch (error) {
    return res.status(500).json({ error: error.message || 'Unable to save registration' });
  }
}
