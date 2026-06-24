const SPREADSHEET_ID = '1r5dTR57jOkQtrdE_OeaAL5IGqH4ZgF-PxsJFZ5vfWyA';
const RECORD_SHEET = '앱기록';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'bootstrap';
  if (action !== 'bootstrap') return json_({ ok: false, error: 'Unknown action' });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return json_({
    ok: true,
    exercises: sheetObjects_(ss.getSheetByName('운동목록')).filter(row => row['사용여부'] === 'Y'),
    routine: sheetObjects_(ss.getSheetByName('주간루틴')).filter(row => row['사용여부'] === 'Y'),
    records: sheetObjects_(ensureRecordSheet_(ss)),
  });
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (payload.action !== 'saveRecord') return json_({ ok: false, error: 'Unknown action' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ensureRecordSheet_(ss);
    const date = String(payload.date || '');
    if (!date) return json_({ ok: false, error: 'date is required' });

    const row = [
      date,
      payload.condition || '',
      payload.status || '',
      payload.mainExercise || '',
      payload.extraExercise || '',
      payload.cardio || '',
      JSON.stringify(payload.addedWorkouts || []),
      JSON.stringify(payload),
      new Date(),
    ];

    const values = sheet.getDataRange().getValues();
    const existingIndex = values.findIndex((item, index) => index > 0 && String(item[0]) === date);
    if (existingIndex > 0) sheet.getRange(existingIndex + 1, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);

    return json_({ ok: true, date });
  } catch (error) {
    return json_({ ok: false, error: String(error && error.message || error) });
  }
}

function ensureRecordSheet_(ss) {
  let sheet = ss.getSheetByName(RECORD_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RECORD_SHEET);
    sheet.appendRow([
      '날짜', '컨디션', '상태', '실제메인운동', '실제추가운동',
      '실제유산소', '추가운동JSON', '전체기록JSON', '수정일시',
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  return values
    .filter(row => row.some(value => value !== ''))
    .map(row => headers.reduce((object, header, index) => {
      object[header] = row[index];
      return object;
    }, {}));
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
