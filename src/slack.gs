function fetchWorkspaceUsers(token, workspaceLabel) {
  if (!token) {
    throw new Error("Slack token is missing: " + workspaceLabel);
  }

  const users = [];
  let cursor = "";
  let isFirstPage = true;

  do {
    if (!isFirstPage) {
      Utilities.sleep(SLACK_USERS_LIST_PAGE_DELAY_MS);
    }

    const url =
      "https://slack.com/api/users.list?limit=" +
      SLACK_USERS_LIST_LIMIT +
      "&cursor=" +
      encodeURIComponent(cursor);

    const response = fetchSlackUsersPageWithRetry(url, token, workspaceLabel);

    const statusCode = response.getResponseCode();
    const bodyText = response.getContentText();

    if (statusCode !== 200) {
      throw new Error(
        "Slack API HTTP error (" +
          workspaceLabel +
          "): " +
          statusCode +
          " body=" +
          bodyText,
      );
    }

    const payload = JSON.parse(bodyText);
    if (!payload.ok) {
      throw new Error(
        "Slack API error (" + workspaceLabel + "): " + payload.error,
      );
    }

    const members = payload.members || [];
    for (let i = 0; i < members.length; i += 1) {
      const m = members[i];
      if (m.deleted || m.is_bot || m.id === "USLACKBOT") {
        continue;
      }

      const profile = m.profile || {};
      const displayName =
        profile.real_name_normalized ||
        profile.real_name ||
        m.real_name ||
        m.name ||
        "";

      users.push({
        id: m.id || "",
        name: displayName,
        email: (profile.email || "").toLowerCase(),
      });
    }

    cursor =
      (payload.response_metadata && payload.response_metadata.next_cursor) ||
      "";
    isFirstPage = false;
  } while (cursor);

  return users;
}

function fetchSlackUsersPageWithRetry(url, token, workspaceLabel) {
  const maxAttempts = 3;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return UrlFetchApp.fetch(url, {
        method: "get",
        headers: {
          Authorization: "Bearer " + token,
        },
        muteHttpExceptions: true,
      });
    } catch (error) {
      const message = (error && error.message) || String(error);
      const isBandwidthError =
        message.indexOf("帯域幅の上限") !== -1 ||
        message.toLowerCase().indexOf("bandwidth") !== -1;

      if (!isBandwidthError || attempt === maxAttempts) {
        throw new Error(
          "Slack API fetch error (" + workspaceLabel + "): " + message,
        );
      }

      Utilities.sleep(SLACK_USERS_LIST_RETRY_DELAY_MS * attempt);
    }
  }

  throw new Error("Slack API fetch failed: " + workspaceLabel);
}
