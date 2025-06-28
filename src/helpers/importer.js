// helpers/importer.js
export async function importJsonHandler(event, sheet) {
  const file = event.currentTarget.files[0];
  if (!file) return;
  try {
    const raw  = await file.text();
    const text = raw.replace(/^\uFEFF/, "").trim();
    if (!text) throw new Error("Empty file");
    const json = JSON.parse(text);
    if (!Array.isArray(json)) throw new Error("JSON must be an array");
    const refs = json.map(r => ({
      cell:    `${r.cell||""}`,
      keyword: `${r.keyword||""}`,
      sheet:   `${r.sheet||""}`,
      type:    `${r.type||"checks"}`,
      value:   `${r.value||""}`,
      attackName: `${r.attackName||""}`,
      critRange: `${r.critRange||""}`,
      damage:  `${r.damage||""}`,
      attackNameCell: `${r.attackNameCell||""}`,
      critRangeCell: `${r.critRangeCell||""}`,
      damageCell: `${r.damageCell||""}`
    }));
    // sheet.actor.setFlag
    await sheet.actor.setFlag("sheexcel_updated","cellReferences", refs);
  } catch(err) {
    console.error("Import JSON error", err);
    ui.notifications.error("Import failed: "+err.message);
  }
}

// Export references as JSON file
export async function exportJsonHandler(sheet) {
  try {
    const refs = await sheet.actor.getFlag("sheexcel_updated", "cellReferences") || [];
    const json = JSON.stringify(refs, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    saveAs(blob, "JasterSheet.json");
  } catch (err) {
    console.error("Export JSON error", err);
    ui.notifications.error("Export failed: " + err.message);
  }
}