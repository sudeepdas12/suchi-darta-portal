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
      { header: 'Submitted At', key: 'submittedAt', width: 25 },
      { header: 'PAN File Name', key: 'panOriginalName', width: 40 },
      { header: 'PAN File Path', key: 'panStoredPath', width: 60 },
      { header: 'Registration File Name', key: 'registrationOriginalName', width: 40 },
      { header: 'Registration File Path', key: 'registrationStoredPath', width: 60 },
      { header: 'Tax Clearance File Name', key: 'taxClearanceOriginalName', width: 40 },
      { header: 'Tax Clearance File Path', key: 'taxClearanceStoredPath', width: 60 }
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
          const rowData = {
            companyName: sheet.getCell('B2').value || '',
            fullName: sheet.getCell('B3').value || '',
            panNumber: sheet.getCell('B4').value || '',
            contactNumber: sheet.getCell('B5').value || '',
            address: sheet.getCell('B6').value || '',
            businessDescription: sheet.getCell('B7').value || '',
            submittedAt: sheet.getCell('B8').value || '',
            panOriginalName: sheet.getCell('B11').value || '',
            panStoredPath: sheet.getCell('B12').value || '',
            registrationOriginalName: sheet.getCell('B15').value || '',
            registrationStoredPath: sheet.getCell('B16').value || '',
            taxClearanceOriginalName: sheet.getCell('B19').value || '',
            taxClearanceStoredPath: sheet.getCell('B20').value || ''
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
