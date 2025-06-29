//import helper for ID-s
import { randomID } from "../helpers/idGenerator.js"; // Make sure this is exported

function assignIdsRecursively(check) {
  check.id = check.id || randomID();
  if (Array.isArray(check.subchecks)) {
    check.subchecks = check.subchecks.map(assignIdsRecursively);
  } else {
    check.subchecks = [];
  }
  return check;
}

// Import JSON file and save references to the actor
export async function importJsonHandler(event, sheet) {
  const file = event.currentTarget.files[0];
  if (!file) return;
  try {
    const raw  = await file.text();
    const text = raw.replace(/^\uFEFF/, "").trim();
    if (!text) throw new Error("Empty file");
    const json = JSON.parse(text);
    if (!Array.isArray(json)) throw new Error("JSON must be an array");
    const refs = json.map(r => assignIdsRecursively({
      ...r,
      cell:    `${r.cell||""}`,
      keyword: `${r.keyword||""}`,
      sheet:   `${r.sheet||""}`,
      type:    `${r.type||"checks"}`,
      value:   `${r.value||""}`,
      attackNameCell: `${r.attackNameCell||""}`,
      critRangeCell: `${r.critRangeCell||""}`,
      damageCell: `${r.damageCell||""}`,
      subchecks: Array.isArray(r.subchecks) ? r.subchecks : []
    }));
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