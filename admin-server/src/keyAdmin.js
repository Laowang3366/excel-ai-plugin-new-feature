export function normalizeKeyIds(ids) {
  if (!Array.isArray(ids)) return [];

  return [...new Set(
    ids
      .map((id) => Math.trunc(Number(id)))
      .filter((id) => Number.isInteger(id) && id > 0)
  )];
}

export function buildExportFilter(filter) {
  if (filter === "unused") {
    return {
      label: "unused",
      where: "WHERE status = 'active' AND used_count = 0",
    };
  }

  return {
    label: "active",
    where: "WHERE status = 'active'",
  };
}
