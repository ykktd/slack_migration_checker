function normalizeName(name) {
  const raw = (name || "").toString();

  const noSpaces = raw.replace(/[\s\u3000]+/g, "");
  const nfkc = noSpaces.normalize("NFKC");
  const removedAlphabet = nfkc.replace(/[A-Za-zＡ-Ｚａ-ｚ]/g, "");

  if (removedAlphabet.length === 0) {
    return nfkc.toLowerCase();
  }

  return removedAlphabet;
}
