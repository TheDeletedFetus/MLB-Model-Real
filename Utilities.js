function MLB_API(url) {
  const response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText());
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  return sheet;
}