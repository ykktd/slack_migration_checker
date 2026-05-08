function buildNewWorkspaceNameMap(newUsers, excludedEmailSet) {
  const map = new Map();

  for (let i = 0; i < newUsers.length; i += 1) {
    const user = newUsers[i];
    if (excludedEmailSet.has((user.email || "").toLowerCase())) {
      continue;
    }

    const key = normalizeName(user.name);
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key).push(user);
  }

  return map;
}

function buildMigrationResultRows(oldUsers, newUsersByName, nameMapping) {
  const rows = [];

  for (let i = 0; i < oldUsers.length; i += 1) {
    const oldUser = oldUsers[i];

    // 修正ルールがあれば適用
    const correctedName =
      nameMapping && nameMapping.has(oldUser.name)
        ? nameMapping.get(oldUser.name)
        : oldUser.name;

    const key = normalizeName(correctedName);
    const matchedUsers = newUsersByName.get(key) || [];

    if (matchedUsers.length === 0) {
      rows.push([STATUS_NOT_MIGRATED, oldUser.name, oldUser.email, "", ""]);
      continue;
    }

    const newUser = matchedUsers.shift();
    if (matchedUsers.length === 0) {
      newUsersByName.delete(key);
    }

    const status = (newUser.email || "").endsWith(MAIL_RULE_SUFFIX)
      ? STATUS_OK
      : STATUS_BAD_EMAIL;

    rows.push([
      status,
      oldUser.name,
      oldUser.email,
      newUser.name,
      newUser.email,
    ]);
  }

  const restKeys = Array.from(newUsersByName.keys());
  for (let i = 0; i < restKeys.length; i += 1) {
    const key = restKeys[i];
    const users = newUsersByName.get(key) || [];

    for (let j = 0; j < users.length; j += 1) {
      const user = users[j];
      rows.push([STATUS_UNMATCHED_NEW_ONLY, "", "", user.name, user.email]);
    }
  }

  rows.sort(function (left, right) {
    const leftStatus = left[0] || "";
    const rightStatus = right[0] || "";
    const leftOrder = Object.prototype.hasOwnProperty.call(
      RESULT_STATUS_ORDER,
      leftStatus,
    )
      ? RESULT_STATUS_ORDER[leftStatus]
      : Number.MAX_SAFE_INTEGER;
    const rightOrder = Object.prototype.hasOwnProperty.call(
      RESULT_STATUS_ORDER,
      rightStatus,
    )
      ? RESULT_STATUS_ORDER[rightStatus]
      : Number.MAX_SAFE_INTEGER;

    if (leftOrder !== rightOrder) {
      return leftOrder - rightOrder;
    }

    const leftOldName = (left[1] || "").toString();
    const rightOldName = (right[1] || "").toString();
    if (leftOldName !== rightOldName) {
      return leftOldName.localeCompare(rightOldName, "ja");
    }

    const leftNewName = (left[3] || "").toString();
    const rightNewName = (right[3] || "").toString();
    if (leftNewName !== rightNewName) {
      return leftNewName.localeCompare(rightNewName, "ja");
    }

    return ((left[4] || "") + "").localeCompare((right[4] || "") + "", "ja");
  });

  return rows;
}
