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
      { header: 'Company Name', key: 'companyName', width: 30 },
      { header: 'Full Name', key: 'fullName', width: 30 },
      { header: 'PAN Number', key: 'panNumber', width: 20 },
      { header: 'Contact Number', key: 'contactNumber', width: 20 },
      { header: 'Address', key: 'address', width: 40 },
      { header: 'Business Description', key: 'businessDescription', width: 40 },
      { header: 'Submitted At', key: 'submittedAt', width: 25 },
      { header: 'PAN File Name', key: 'panOriginalName', width: 40 },
      { header: 'PAN Stored Path', key: 'panStoredPath', width: 60 },
      { header: 'PAN File ID', key: 'panFileId', width: 36 },
      { header: 'Registration File Name', key: 'registrationOriginalName', width: 40 },
      { header: 'Registration Stored Path', key: 'registrationStoredPath', width: 60 },
      { header: 'Registration File ID', key: 'registrationFileId', width: 36 },
      { header: 'Tax Clearance File Name', key: 'taxClearanceOriginalName', width: 40 },
      { header: 'Tax Clearance Stored Path', key: 'taxClearanceStoredPath', width: 60 },
      { header: 'Tax Clearance File ID', key: 'taxClearanceFileId', width: 36 }
    ];

    sheet.addRow({
      companyName: safeText(payload.companyName),
      fullName: safeText(payload.fullName),
      panNumber: safeText(payload.panNumber),
      contactNumber: safeText(payload.contactNumber),
      address: safeText(payload.address),
      businessDescription: safeText(payload.businessDescription),
      submittedAt: new Date().toISOString(),
      panOriginalName: safeText(payload.documents?.panOriginalName),
      panStoredPath: safeText(payload.documents?.panStoredPath),
      panFileId: safeText(payload.documents?.panFileId),
      registrationOriginalName: safeText(payload.documents?.registrationOriginalName),
      registrationStoredPath: safeText(payload.documents?.registrationStoredPath),
      registrationFileId: safeText(payload.documents?.registrationFileId),
      taxClearanceOriginalName: safeText(payload.documents?.taxClearanceOriginalName),
      taxClearanceStoredPath: safeText(payload.documents?.taxClearanceStoredPath),
      taxClearanceFileId: safeText(payload.documents?.taxClearanceFileId)
    });

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
