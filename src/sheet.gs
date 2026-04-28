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

  sheet.clear();

  const output = [RESULT_HEADER].concat(resultRows);
  sheet.getRange(1, 1, output.length, RESULT_HEADER.length).setValues(output);

  applyMigrationResultSheetFormatting(
    sheet,
    output.length,
    RESULT_HEADER.length,
  );
}

function applyMigrationResultSheetFormatting(sheet, rowCount, columnCount) {
  if (rowCount <= 0 || columnCount <= 0) {
    return;
  }

  sheet.setFrozenRows(1);
  sheet
    .getRange(1, 1, 1, columnCount)
    .setFontWeight("bold")
    .setBackground("#1f4e78")
    .setFontColor("#ffffff");

  if (rowCount > 1) {
    sheet
      .getRange(2, 1, rowCount - 1, columnCount)
      .setWrap(true)
      .setVerticalAlignment("middle");

    const bodyBackgrounds = [];
    for (let rowIndex = 1; rowIndex < rowCount; rowIndex += 1) {
      const status = sheet.getRange(rowIndex + 1, 1).getValue();
      const backgroundColor = RESULT_STATUS_COLORS[status] || "#ffffff";
      const rowBackground = [];
      for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
        rowBackground.push(backgroundColor);
      }
      bodyBackgrounds.push(rowBackground);
    }

    sheet
      .getRange(2, 1, rowCount - 1, columnCount)
      .setBackgrounds(bodyBackgrounds);
  }

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 220);

  const dataRange = sheet.getRange(1, 1, rowCount, columnCount);
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  dataRange.createFilter();
}

function getNameMappingRules(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    return new Map();
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return new Map();
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const mapping = new Map();

  for (let i = 0; i < values.length; i += 1) {
    const oldName = (values[i][0] || "").toString().trim();
    const newName = (values[i][1] || "").toString().trim();

    if (oldName && newName) {
      mapping.set(oldName, newName);
    }
  }

  return mapping;
}

function ensureNameMappingSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet
      .getRange(1, 1, 1, NAME_MAPPING_HEADER.length)
      .setValues([NAME_MAPPING_HEADER])
      .setFontWeight("bold")
      .setBackground("#1f4e78")
      .setFontColor("#ffffff");

    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 220);
  }

  return sheet;
}
