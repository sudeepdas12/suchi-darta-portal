import B2 from 'backblaze-b2';
import ExcelJS from 'exceljs';

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

    // For each summary file, download and extract document metadata (fileIds, original names, stored paths)
    const detailed = [];
    for (const sf of summaryFiles) {
      try {
        const downloadResp = await b2.downloadFileById({ fileId: sf.fileId, responseType: 'arraybuffer' });
        const buffer = Buffer.from(downloadResp.data);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const sheet = workbook.getWorksheet('Registration');
        if (sheet) {
          const row = sheet.getRow(2);
          detailed.push({
            fileName: sf.fileName,
            fileId: sf.fileId,
            uploadTimestamp: sf.uploadTimestamp,
            companyName: row.getCell(1).value || '',
            fullName: row.getCell(2).value || '',
            panNumber: row.getCell(3).value || '',
            contactNumber: row.getCell(4).value || '',
            address: row.getCell(5).value || '',
            businessDescription: row.getCell(6).value || '',
            shopType: row.getCell(7).value || '',
            submittedAt: row.getCell(8).value || '',
            documents: {
              panOriginalName: row.getCell(9).value || '',
              panStoredPath: row.getCell(10).value || '',
              panFileId: row.getCell(11).value || '',
              registrationOriginalName: row.getCell(12).value || '',
              registrationStoredPath: row.getCell(13).value || '',
              registrationFileId: row.getCell(14).value || '',
              taxClearanceOriginalName: row.getCell(15).value || '',
              taxClearanceStoredPath: row.getCell(16).value || '',
              taxClearanceFileId: row.getCell(17).value || '',
              rateOriginalName: row.getCell(18).value || '',
              rateStoredPath: row.getCell(19).value || '',
              rateFileId: row.getCell(20).value || '',
              othersOriginalName: row.getCell(21).value || '',
              othersStoredPath: row.getCell(22).value || '',
              othersFileId: row.getCell(23).value || ''
            }
          });
        } else {
          detailed.push({ ...sf, documents: {} });
        }
      } catch (err) {
        console.error('Failed to parse summary', sf.fileName, err.message);
        detailed.push({ ...sf, documents: {} });
      }
    }

    return res.status(200).json({ summaries: detailed });
  } catch (error) {
    return res.status(500).json({ error: error.message || 'Unable to list summary files' });
  }
}
