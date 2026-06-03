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

    // List all files
    let allFiles = [];
    let nextFileName;

    do {
      const response = await b2.listFileNames({
        bucketId: B2_BUCKET_ID,
        startFileName: nextFileName,
        maxFileCount: 100
      });

      const batch = response.data.files || [];
      allFiles.push(...batch);
      nextFileName = response.data.nextFileName;
    } while (nextFileName);

    // Filter summary files
    const summaryFiles = allFiles.filter(f => 
      f.fileName.toLowerCase().endsWith('.xlsx') && 
      f.fileName.includes('registration-')
    );

    // Create master workbook
    const masterWorkbook = new ExcelJS.Workbook();
    const masterSheet = masterWorkbook.addWorksheet('All Submissions');

    masterSheet.columns = [
      { header: 'Company Name', key: 'companyName', width: 30 },
      { header: 'Full Name', key: 'fullName', width: 30 },
      { header: 'PAN Number', key: 'panNumber', width: 20 },
      { header: 'Contact Number', key: 'contactNumber', width: 20 },
      { header: 'Address', key: 'address', width: 40 },
      { header: 'Business Description', key: 'businessDescription', width: 40 },
      { header: 'Shop Type', key: 'shopType', width: 20 },
      { header: 'Submitted At', key: 'submittedAt', width: 25 },
      { header: 'PAN File Name', key: 'panOriginalName', width: 40 },
      { header: 'PAN File Path', key: 'panStoredPath', width: 60 },
      { header: 'PAN File ID', key: 'panFileId', width: 36 },
      { header: 'Registration File Name', key: 'registrationOriginalName', width: 40 },
      { header: 'Registration File Path', key: 'registrationStoredPath', width: 60 },
      { header: 'Registration File ID', key: 'registrationFileId', width: 36 },
      { header: 'Tax Clearance File Name', key: 'taxClearanceOriginalName', width: 40 },
      { header: 'Tax Clearance File Path', key: 'taxClearanceStoredPath', width: 60 },
      { header: 'Tax Clearance File ID', key: 'taxClearanceFileId', width: 36 },
      { header: 'Others File Name', key: 'othersOriginalName', width: 40 },
      { header: 'Others File Path', key: 'othersStoredPath', width: 60 },
      { header: 'Others File ID', key: 'othersFileId', width: 36 }
    ];

    // Download and extract data from each summary file
    for (const summaryFile of summaryFiles) {
      try {
        const downloadResponse = await b2.downloadFileById({
          fileId: summaryFile.fileId,
          responseType: 'arraybuffer'
        });

        const fileBuffer = Buffer.from(downloadResponse.data);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(fileBuffer);
        const sheet = workbook.getWorksheet('Registration');

        if (sheet) {
          // Get the first data row (row 2, since row 1 is headers)
          const row = sheet.getRow(2);
          const rowData = {
            companyName: row.getCell(1).value || '',
            fullName: row.getCell(2).value || '',
            panNumber: row.getCell(3).value || '',
            contactNumber: row.getCell(4).value || '',
            address: row.getCell(5).value || '',
            businessDescription: row.getCell(6).value || '',
            shopType: row.getCell(7).value || '',
            submittedAt: row.getCell(8).value || '',
            panOriginalName: row.getCell(9).value || '',
            panStoredPath: row.getCell(10).value || '',
            panFileId: row.getCell(11).value || '',
            registrationOriginalName: row.getCell(12).value || '',
            registrationStoredPath: row.getCell(13).value || '',
            registrationFileId: row.getCell(14).value || '',
            taxClearanceOriginalName: row.getCell(15).value || '',
            taxClearanceStoredPath: row.getCell(16).value || '',
            taxClearanceFileId: row.getCell(17).value || '',
            othersOriginalName: row.getCell(18).value || '',
            othersStoredPath: row.getCell(19).value || '',
            othersFileId: row.getCell(20).value || ''
          };

          masterSheet.addRow(rowData);
        }
      } catch (error) {
        console.error(`Failed to process ${summaryFile.fileName}:`, error.message);
      }
    }

    // Generate buffer
    const buffer = await masterWorkbook.xlsx.writeBuffer();

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="consolidated-submissions-${Date.now()}.xlsx"`);
    return res.status(200).send(buffer);
  } catch (error) {
    return res.status(500).json({ error: error.message || 'Unable to generate consolidated report' });
  }
}
