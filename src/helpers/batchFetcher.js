// helpers/batchFetcher.js

/**
 * Given a Google Sheets ID and an array of reference objects, 
 * batch-fetches the requested cell values (including extra attack fields)
 * and writes them back into each ref before returning the updated array.
 * @param {string} sheetId        The Sheets spreadsheet ID
 * @param {Array<Object>} refs     Array of refs, each with:
 *    - sheet, cell, type, attackNameCell, critRangeCell, damageCell
 * @returns {Promise<Array<Object>>}  The same refs array, with .value, .attackName, .critRange, .damage populated
 */
export async function batchFetchValues(sheetId, refs) {
  const apiKey = "AIzaSyCAYacdw4aB7GtoxwnlpaF3aFZ2DgcJNHo";

  // Build a list of queries, one per referenced cell
  const queries = refs.flatMap((r, idx) => {
    const sheetName = r.sheet.match(/[^A-Za-z0-9_]/)
      ? `'${r.sheet.replace(/'/g, "''")}'`
      : r.sheet;
    const list = [];

    // Core value cell
    if (r.cell) {
      list.push({ idx, field: "value", range: `${sheetName}!${r.cell}` });
    }

    // Attack-specific fields
    if (r.type === "attacks") {
      if (r.attackNameCell) {
        list.push({ idx, field: "attackName", range: `${sheetName}!${r.attackNameCell}` });
      }
      if (r.critRangeCell) {
        list.push({ idx, field: "critRange", range: `${sheetName}!${r.critRangeCell}` });
      }
      if (r.damageCell) {
        list.push({ idx, field: "damage", range: `${sheetName}!${r.damageCell}` });
      }
    }

    return list;
  });

  if (!queries.length) return refs;

  // Construct the batchGet URL
  const rangesParam = queries.map(q => `ranges=${encodeURIComponent(q.range)}`).join("&");
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchGet?key=${apiKey}&${rangesParam}`;

  const res = await fetch(url);
  if (!res.ok) throw new Error(`Google Sheets API error: ${res.status} ${res.statusText}`);
  const json = await res.json();

  // Apply returned values back onto refs
  json.valueRanges.forEach((vr, i) => {
    const { idx, field } = queries[i];
    refs[idx][field] = vr.values?.[0]?.[0] ?? "";
  });

  return refs;
}
