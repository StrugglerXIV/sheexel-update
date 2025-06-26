// helpers/prepareData.js

/**
 * Enriches the Foundry-generated sheet data with Sheexcel flags and grouping.
 * @param {Object} data   The object returned by super.getData()
 * @param {Actor}  actor  The Actor document
 * @returns {Object}      The same data object, mutated with extra fields
 */
export async function prepareSheetData(data, actor) {
  // Pull in your saved flags from the actor
  data.sheetUrl         = actor.getFlag("sheexcel_updated", "sheetUrl")       || "";
  data.hideMenu         = actor.getFlag("sheexcel_updated", "hideMenu")       || false;
  data.zoomLevel        = actor.getFlag("sheexcel_updated", "zoomLevel")     ?? 100;
  data.sidebarCollapsed = actor.getFlag("sheexcel_updated", "sidebarCollapsed") || false;

  // Read refs & sheetNames from the actor
  const refs       = actor.getFlag("sheexcel_updated", "cellReferences") || [];
  const sheetNames = actor.getFlag("sheexcel_updated", "sheetNames")     || [];

  // Build adjustedReferences
  data.adjustedReferences = refs.map((r, i) => ({
    ...r,
    index:            i,
    sheetNames,
    // Attack-specific
    attackNameCell:   r.attackNameCell || "",
    critRangeCell:    r.critRangeCell  || "",
    damageCell:       r.damageCell     || ""
  }));

  // Group into checks/saves/attacks/spells
  data.groupedReferences = { checks: [], saves: [], attacks: [], spells: [] };
  data.adjustedReferences.forEach(r => {
    (data.groupedReferences[r.type] || data.groupedReferences.checks).push(r);
  });

  return data;
}
