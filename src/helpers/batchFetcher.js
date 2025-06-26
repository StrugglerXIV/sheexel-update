// helpers/batchFetcher.js
export async function batchFetchValues(sheetId, refs) {
  const apiKey = "AIzaSyCAYacdw4aB7GtoxwnlpaF3aFZ2DgcJNHo";
  const queries = refs.map((r,i) => {
    const safe = /[^A-Za-z0-9_]/.test(r.sheet)
      ? `'${r.sheet.replace(/'/g,"''")}'!${r.cell}`
      : `${r.sheet}!${r.cell}`;
    return { idx:i, range: safe };
  }).filter(q => q.range.includes("!"));

  if (queries.length) {
    const params = queries.map(q => `ranges=${encodeURIComponent(q.range)}`).join("&");
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchGet?key=${apiKey}&${params}`;
    try {
      const res  = await fetch(url);
      if (!res.ok) throw new Error(res.statusText);
      const json = await res.json();
      json.valueRanges.forEach((vr,i) => {
        refs[queries[i].idx].value = vr.values?.[0]?.[0] || "";
      });
    } catch(err) {
      console.warn("Batch fetch failed", err);
    }
  }

  return refs;
}
