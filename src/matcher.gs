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

function buildMigrationResultRows(oldUsers, newUsersByName) {
  const rows = [];

  for (let i = 0; i < oldUsers.length; i += 1) {
    const oldUser = oldUsers[i];
    const key = normalizeName(oldUser.name);
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

  return rows;
}
