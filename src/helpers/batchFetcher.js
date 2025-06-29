export async function batchFetchValues(sheetId, refs) {
  const apiKey = game.settings.get("sheexcel_updated", "googleApiKey");

  // Helper to recursively collect all queries with a path
  function collectQueries(refs, path = []) {
    let queries = [];
    refs.forEach((r, idx) => {
      const sheetName = r.sheet.match(/[^A-Za-z0-9_]/)
        ? `'${r.sheet.replace(/'/g, "''")}'`
        : r.sheet;

      // Main value
      if (r.cell) {
        queries.push({ path: [...path, idx], field: "value", range: `${sheetName}!${r.cell}` });
      }
      // Attack-specific fields
      if (r.type === "attacks") {
        if (r.attackNameCell) {
          queries.push({ path: [...path, idx], field: "attackName", range: `${sheetName}!${r.attackNameCell}` });
        }
        if (r.critRangeCell) {
          queries.push({ path: [...path, idx], field: "critRange", range: `${sheetName}!${r.critRangeCell}` });
        }
        if (r.damageCell) {
          queries.push({ path: [...path, idx], field: "damage", range: `${sheetName}!${r.damageCell}` });
        }
      }
      // Recurse into subchecks
      if (Array.isArray(r.subchecks) && r.subchecks.length) {
        queries = queries.concat(collectQueries(r.subchecks, [...path, idx, "subchecks"]));
      }
    });
    return queries;
  }

  // Helper to assign a value back into the nested structure by path
  function assignByPath(obj, path, field, value) {
    let ref = obj;
    for (let i = 0; i < path.length - 1; i++) {
      ref = ref[path[i]];
    }
    ref[path[path.length - 1]][field] = value;
  }

  // Collect all queries
  const queries = collectQueries(refs);

  if (!queries.length) return refs;

  // Build the batchGet URL for all ranges
  const rangesParam = queries.map(q => `ranges=${encodeURIComponent(q.range)}`).join("&");
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchGet?key=${apiKey}&${rangesParam}`;

  // Fetch all requested cell values in a single API call
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Google Sheets API error: ${res.status} ${res.statusText}`);
  const json = await res.json();

  // Assign the fetched values back to the correct fields in each ref/subcheck
  json.valueRanges.forEach((vr, i) => {
    const { path, field } = queries[i];
    assignByPath(refs, path, field, vr.values?.[0]?.[0] ?? "");
  });

  return refs;
}