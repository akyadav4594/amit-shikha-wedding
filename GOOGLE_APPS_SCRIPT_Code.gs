const SHEET_NAME = 'Sheet1';

function doGet() {
  return json_({ ok: true, message: 'Amit & Shikha RSVP endpoint is live.' });
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet '${SHEET_NAME}' was not found.`);

    // The wedding website sends URL-encoded form data.
    // Support that format, plus JSON for easier future testing.
    let data = {};
    if (e && e.postData && e.postData.contents) {
      const raw = e.postData.contents;
      const type = (e.postData.type || '').toLowerCase();
      if (type.indexOf('application/json') !== -1) {
        data = JSON.parse(raw || '{}');
      } else {
        data = e.parameter || {};
      }
    } else {
      data = (e && e.parameter) ? e.parameter : {};
    }

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

    return json_({ success: true, message: 'RSVP saved.' });
  } catch (error) {
    return json_({ success: false, error: error.message });
  }
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
