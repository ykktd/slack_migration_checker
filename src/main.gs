function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Slack移行チェック")
    .addItem("移行チェックを実行", "menuRunMigrationCheck")
    .addToUi();
}

function runMigrationCheck() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const oldToken = requireScriptProperty(scriptProperties, "SLACK_TOKEN_OLD");
  const newToken = requireScriptProperty(scriptProperties, "SLACK_TOKEN_NEW");

  // 修正ルールシートを初期化（存在しなければ作成）
  ensureNameMappingSheet(SHEET_NAME_NAME_MAPPING);

  const excludedEmailSet = getExcludedEmailSet(SHEET_NAME_EXCLUSION);
  const nameMapping = getNameMappingRules(SHEET_NAME_NAME_MAPPING);
  const oldUsers = fetchWorkspaceUsers(oldToken, "old");
  const newUsers = fetchWorkspaceUsers(newToken, "new");

  const filteredOldUsers = oldUsers.filter(
    (user) => !excludedEmailSet.has((user.email || "").toLowerCase()),
  );
  const newUsersByName = buildNewWorkspaceNameMap(newUsers, excludedEmailSet);
  const rows = buildMigrationResultRows(
    filteredOldUsers,
    newUsersByName,
    nameMapping,
  );

  writeMigrationResults(rows, SHEET_NAME_RESULT);
  Logger.log("runMigrationCheck finished. rowCount=" + rows.length);
}

function exportCurrentNewUsersToExclusionList() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const newToken = requireScriptProperty(scriptProperties, "SLACK_TOKEN_NEW");

  const newUsers = fetchWorkspaceUsers(newToken, "new");
  const writtenCount = appendUsersToExclusionList(
    newUsers,
    SHEET_NAME_EXCLUSION,
  );

  Logger.log(
    "exportCurrentNewUsersToExclusionList finished. writtenCount=" +
      writtenCount,
  );
}

function menuExportExclusionList() {
  exportCurrentNewUsersToExclusionList();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "除外リストの更新が完了しました",
    "Slack移行チェック",
    5,
  );
}

function menuRunMigrationCheck() {
  runMigrationCheck();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "移行チェックが完了しました。移行確認結果シートを確認してください",
    "Slack移行チェック",
    5,
  );
}

function requireScriptProperty(scriptProperties, key) {
  const value = scriptProperties.getProperty(key);
  if (!value) {
    throw new Error("Script property is missing: " + key);
  }
  return value;
}
