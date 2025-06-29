// --- Helpers ---
export function randomID() {
  return Math.random().toString(36).substr(2, 9);
}

export function findExistingByIdOrKeyword(existingRefs, id, keyword) {
  for (const ref of existingRefs) {
    if (ref.id === id || ref.keyword === keyword) return ref;
    if (Array.isArray(ref.subchecks)) {
      const found = findExistingByIdOrKeyword(ref.subchecks, id, keyword);
      if (found) return found;
    }
  }
  return null;
}