function fetchWorkspaceUsers(token, workspaceLabel) {
  if (!token) {
    throw new Error("Slack token is missing: " + workspaceLabel);
  }

  const users = [];
  let cursor = "";

  do {
    const url =
      "https://slack.com/api/users.list?limit=200&cursor=" +
      encodeURIComponent(cursor);
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        Authorization: "Bearer " + token,
      },
      muteHttpExceptions: true,
    });

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
  } while (cursor);

  return users;
}
