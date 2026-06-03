import B2 from 'backblaze-b2';
import ExcelJS from 'exceljs';

function safeText(value) {
  return value == null ? '' : String(value);
}

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

  const userFolder = safeText(payload.userFolder).replace(/[^a-zA-Z0-9._-]+/g, '_') || 'unknown-user';
  const timestamp = Date.now();
  const excelFileName = `${userFolder}/registration-${timestamp}.xlsx`;

  const b2 = new B2({
    applicationKeyId: B2_KEY_ID,
    applicationKey: B2_APP_KEY
  });

  try {
    await b2.authorize();

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Registration');
    sheet.columns = [
      { header: 'Field', key: 'field', width: 30 },
      { header: 'Value', key: 'value', width: 70 }
    ];

    sheet.addRow({ field: 'Company Name', value: safeText(payload.companyName) });
    sheet.addRow({ field: 'Full Name', value: safeText(payload.fullName) });
    sheet.addRow({ field: 'PAN Number', value: safeText(payload.panNumber) });
    sheet.addRow({ field: 'Contact Number', value: safeText(payload.contactNumber) });
    sheet.addRow({ field: 'Address', value: safeText(payload.address) });
    sheet.addRow({ field: 'Business Description', value: safeText(payload.businessDescription) });
    sheet.addRow({ field: 'Submitted At', value: new Date().toISOString() });
    sheet.addRow({ field: 'User Folder', value: userFolder });

    sheet.addRow({ field: '', value: '' });
    sheet.addRow({ field: 'Attachment', value: 'Details' });
    sheet.addRow({ field: 'PAN File Name', value: safeText(payload.documents?.panOriginalName) });
    sheet.addRow({ field: 'PAN Stored Path', value: safeText(payload.documents?.panStoredPath) });
    sheet.addRow({ field: 'PAN File ID', value: safeText(payload.documents?.panFileId) });
    sheet.addRow({ field: '' , value: '' });
    sheet.addRow({ field: 'Registration File Name', value: safeText(payload.documents?.registrationOriginalName) });
    sheet.addRow({ field: 'Registration Stored Path', value: safeText(payload.documents?.registrationStoredPath) });
    sheet.addRow({ field: 'Registration File ID', value: safeText(payload.documents?.registrationFileId) });
    sheet.addRow({ field: '' , value: '' });
    sheet.addRow({ field: 'Tax Clearance File Name', value: safeText(payload.documents?.taxClearanceOriginalName) });
    sheet.addRow({ field: 'Tax Clearance Stored Path', value: safeText(payload.documents?.taxClearanceStoredPath) });
    sheet.addRow({ field: 'Tax Clearance File ID', value: safeText(payload.documents?.taxClearanceFileId) });

    const buffer = await workbook.xlsx.writeBuffer();

    const uploadUrlResponse = await b2.getUploadUrl({ bucketId: B2_BUCKET_ID });

    const uploadResponse = await fetch(uploadUrlResponse.data.uploadUrl, {
      method: 'POST',
      headers: {
        Authorization: uploadUrlResponse.data.authorizationToken,
        'X-Bz-File-Name': encodeURIComponent(excelFileName),
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'X-Bz-Content-Sha1': 'do_not_verify'
      },
      body: buffer
    });

    if (!uploadResponse.ok) {
      const errorText = await uploadResponse.text().catch(() => 'Unknown upload error');
      throw new Error(`Backblaze upload failed: ${uploadResponse.status} ${errorText}`);
    }

    const uploadResult = await uploadResponse.json();
    return res.status(200).json({ success: true, excelFileName, excelFileId: uploadResult.fileId });
  } catch (error) {
    return res.status(500).json({ error: error.message || 'Unable to save registration' });
  }
}
