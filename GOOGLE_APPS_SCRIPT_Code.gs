const SHEET_NAME = 'Sheet1';

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet '${SHEET_NAME}' was not found.`);

    const data = JSON.parse(e.postData.contents || '{}');
    const attendance = data.attendance === 'Yes' ? 'Yes' : 'No';
    const nov20 = attendance === 'Yes' && data.nov20 === '20Nov' ? '20Nov' : '';
    const nov21 = attendance === 'Yes' && data.nov21 === '21Nov' ? '21Nov' : '';

    sheet.appendRow([
      new Date(),
      data.guest || '',
      data.phone || '',
      data.guests || '',
      attendance,
      nov20,
      nov21,
      data.message || ''
    ]);

    return json_({ success: true });
  } catch (error) {
    return json_({ success: false, error: error.message });
  }
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
