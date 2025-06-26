// helpers/prepareData.js

/**
 * Enriches the Foundry-generated sheet data with Sheexcel flags and grouping.
 * @param {Object} data   The object returned by super.getData()
 * @param {Actor}  actor  The Actor document
 * @returns {Object}      The same data object, mutated with extra fields
 */
export function prepareSheetData(data, actor) {
  // Pull in your saved flags
  data.sheetUrl         = actor.getFlag("sheexcel_updated","sheetUrl")       || "";
  data.hideMenu         = actor.getFlag("sheexcel_updated","hideMenu")       || false;
  data.zoomLevel        = actor.getFlag("sheexcel_updated","zoomLevel")     ?? 100;
  data.sidebarCollapsed = actor.getFlag("sheexcel_updated","sidebarCollapsed") || false;

  // Build your references array
  const refs   = actor.getFlag("sheexcel_updated","cellReferences") || [];
  const sheets = actor.getFlag("sheexcel_updated","sheetNames")     || [];
  data.adjustedReferences = refs.map((r,i) => ({ 
    ...r, 
    index:      i, 
    sheetNames: sheets 
  }));

  // Group them for the Main tab
  data.groupedReferences = { checks:[], saves:[], attacks:[], spells:[] };
  refs.forEach((r,i) => {
    const type = r.type || "checks";
    if ( data.groupedReferences[type] ) {
      data.groupedReferences[type].push({ ...r, index: i });
    }
  });

  return data;
}
