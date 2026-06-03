import B2 from 'backblaze-b2';
import ExcelJS from 'exceljs';
import archiver from 'archiver';

const ADMIN_TOKEN = process.env.ADMIN_TOKEN || '';

function isAuthorized(token) {
  if (!ADMIN_TOKEN) {
    return true;
  }
  return token === ADMIN_TOKEN;
}

function sanitizeFileName(name) {
  return String(name || 'file')
    .trim()
    .replace(/[/\\?%*:|"<>]/g, '-')
    .replace(/\s+/g, '_')
    .substring(0, 120);
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

  const summaryFileId = req.query.fileId;
  if (!summaryFileId) {
    return res.status(400).json({ error: 'Missing fileId query parameter' });
  }

  const b2 = new B2({
    applicationKeyId: B2_KEY_ID,
    applicationKey: B2_APP_KEY
  });

  try {
    await b2.authorize();

    const downloadResponse = await b2.downloadFileById({
      fileId: summaryFileId,
      responseType: 'arraybuffer'
    });

    const summaryBuffer = Buffer.from(downloadResponse.data);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(summaryBuffer);
    const sheet = workbook.getWorksheet('Registration');
    const row = sheet ? sheet.getRow(2) : null;

    const companyName = row ? String(row.getCell(1).value || 'vendor').trim() : 'vendor';
    const panNumber = row ? String(row.getCell(3).value || 'unknown-pan').trim() : 'unknown-pan';
    const archiveName = `vendor-${sanitizeFileName(companyName)}-${sanitizeFileName(panNumber)}.zip`;

    const attachments = [];
    if (row) {
      const summaryName = `registration-summary-${sanitizeFileName(companyName)}.xlsx`;
      attachments.push({ buffer: summaryBuffer, name: summaryName });

      const documentEntries = [
        { key: 'panFileId', label: 'PAN' },
        { key: 'registrationFileId', label: 'CompanyRegistration' },
        { key: 'taxClearanceFileId', label: 'TaxClearance' },
        { key: 'rateFileId', label: 'RateOfProducts' },
        { key: 'othersFileId', label: 'Others' }
      ];

      documentEntries.forEach((entry, index) => {
        const fileIdCellIndex = 11 + index * 4;
        const originalNameCellIndex = 9 + index * 4;
        const fileId = String(row.getCell(fileIdCellIndex).value || '').trim();
        const originalName = String(row.getCell(originalNameCellIndex).value || '').trim();
        if (fileId) {
          attachments.push({ fileId, originalName, label: entry.label });
        }
      });
    }

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${archiveName}"`);

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.on('error', err => {
      if (!res.headersSent) {
        res.status(500).json({ error: err.message || 'Archive generation failed' });
      }
      res.end();
    });
    archive.pipe(res);

    for (const attachment of attachments) {
      if (attachment.buffer) {
        archive.append(attachment.buffer, { name: attachment.name });
        continue;
      }

      try {
        const fileResp = await b2.downloadFileById({ fileId: attachment.fileId, responseType: 'arraybuffer' });
        const fileBuffer = Buffer.from(fileResp.data);
        const entryName = sanitizeFileName(attachment.originalName || `${attachment.label}.pdf`);
        archive.append(fileBuffer, { name: `attachments/${entryName}` });
      } catch (err) {
        archive.append(`Failed to download ${attachment.label}: ${err.message}`, { name: `attachments/${attachment.label}-error.txt` });
      }
    }

    await archive.finalize();
  } catch (error) {
    if (!res.headersSent) {
      return res.status(500).json({ error: error.message || 'Unable to create vendor bundle' });
    }
  }
}
