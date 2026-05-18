const SHEET_ID = '1QbWN2TARGN7vifJF5wv9BBu-dXlF6Kp3Ew3qQPiPYpE';
const SHEET_NAME = 'profile';

function doGet(e) {
  const id = e && e.parameter && e.parameter.id ? String(e.parameter.id).trim() : '';
  const employee = id ? getEmployee(id) : null;

  // Format data as a clean JSON package
  const result = {
    employee: employee,
    requestedId: id
  };

  // Return JSON instead of an HTML Template
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getEmployee(id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const empId = String(data[i][0]).trim();
    if (empId === String(id).trim()) {
      return {
        id: data[i][0],
        name: data[i][1],
        position: data[i][2],
        department: data[i][3]
      };
    }
  }
  return null;
}
