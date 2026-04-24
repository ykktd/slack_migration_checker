function getExcludedEmailSet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Sheet not found: " + sheetName);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return new Set();
  }

  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const excluded = new Set();

  for (let i = 0; i < values.length; i += 1) {
    const mail = (values[i][0] || "").toString().trim().toLowerCase();
    if (mail) {
      excluded.add(mail);
    }
  }

  return excluded;
}

function appendUsersToExclusionList(users, sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet
      .getRange(1, 1, 1, EXCLUSION_HEADER.length)
      .setValues([EXCLUSION_HEADER]);
  }

  const currentLastRow = sheet.getLastRow();
  const existingValues =
    currentLastRow > 1
      ? sheet.getRange(2, 2, currentLastRow - 1, 1).getValues()
      : [];

  const existingEmailSet = new Set();
  for (let i = 0; i < existingValues.length; i += 1) {
    const mail = (existingValues[i][0] || "").toString().trim().toLowerCase();
    if (mail) {
      existingEmailSet.add(mail);
    }
  }

  const rowsToAppend = [];
  for (let i = 0; i < users.length; i += 1) {
    const user = users[i];
    const mail = (user.email || "").toString().trim().toLowerCase();
    if (!mail || existingEmailSet.has(mail)) {
      continue;
    }

    rowsToAppend.push([(user.name || "").toString().trim(), mail]);
    existingEmailSet.add(mail);
  }

  if (rowsToAppend.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 2)
      .setValues(rowsToAppend);
  }

  return rowsToAppend.length;
}

function writeMigrationResults(resultRows, sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  sheet.clearContents();

  const output = [RESULT_HEADER].concat(resultRows);
  sheet.getRange(1, 1, output.length, RESULT_HEADER.length).setValues(output);
}
