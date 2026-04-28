const MAIL_RULE_SUFFIX = "17th.fw@gmail.com";
const SLACK_USERS_LIST_LIMIT = 100;
const SLACK_USERS_LIST_PAGE_DELAY_MS = 1000;
const SLACK_USERS_LIST_RETRY_DELAY_MS = 3000;

const SHEET_NAME_EXCLUSION = "除外リスト";
const SHEET_NAME_RESULT = "移行確認結果";

const EXCLUSION_HEADER = ["氏名", "メアド"];

const STATUS_OK = "移行完了（正常）";
const STATUS_BAD_EMAIL = "メアド不正";
const STATUS_NOT_MIGRATED = "未移行";
const STATUS_UNMATCHED_NEW_ONLY = "照合不可（新WSのみ存在）";

const RESULT_HEADER = [
  "判定ステータス",
  "旧WS 氏名",
  "旧WS メアド",
  "新WS 氏名",
  "新WS メアド",
];
