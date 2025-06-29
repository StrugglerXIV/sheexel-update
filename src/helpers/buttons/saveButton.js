// Import the batch fetcher for Google Sheets and a helper to find existing references
import { batchFetchValues } from "../batchFetcher.js";
import { findExistingByIdOrKeyword, randomID } from "../idGenerator.js";

/**
 * Handles saving all reference rows from the sheet UI to the actor's flags.
 * Optionally fetches updated values from Google Sheets if a sheetId is set.
 * 
 * @param {SheexcelActorSheet} sheet - The sheet instance (usually 'this' from the sheet class)
 * @param {Event} event - The triggering event (usually a button click)
 */
export async function onSaveReferences(sheet, event) {
  event.preventDefault();

  const rows = sheet.element.find(".sheexcel-reference-row").toArray();
  const existingRefs = sheet.actor.getFlag("sheexcel_updated", "cellReferences") || [];

  console.log("Saving references. Existing refs:", JSON.parse(JSON.stringify(existingRefs)));
  console.log("Found reference rows in UI:", rows.length);

  const refs = rows.map((el, idx) => {
    const $r = $(el);
    const keyword = $r.find("input[data-type='keyword']").val().trim();
    const id = $r.data("id") || randomID();
    const existing = findExistingByIdOrKeyword(existingRefs, id, keyword);

    // --- Collect subchecks for this row ---
    const subcheckRows = $r.find(".sheexcel-subcheck-row").toArray();
    let subchecks;
    if (subcheckRows.length > 0) {
      subchecks = subcheckRows.map((subEl, subIdx) => {
        const $sub = $(subEl);
        const subKeyword = $sub.find("input[data-type='keyword']").val().trim();
        const subId = $sub.data("id") || randomID();
        const subExisting = findExistingByIdOrKeyword(existing?.subchecks || [], subId, subKeyword);

        const subObj = {
          id: subExisting?.id || subId,
          cell: $sub.find("input[data-type='cell']").val().trim(),
          keyword: subKeyword,
          sheet: $sub.find("select[data-type='sheet']").val(),
          type: $sub.find("select[data-type='refType']").val(),
          attackNameCell: $sub.find("input[data-type='attackNameCell']").val()?.trim() || "",
          critRangeCell: $sub.find("input[data-type='critRangeCell']").val()?.trim() || "",
          damageCell: $sub.find("input[data-type='damageCell']").val()?.trim() || "",
          value: "",
          subchecks: subExisting?.subchecks || []
        };
        console.log(`Subcheck [${idx}][${subIdx}]:`, subObj);
        return subObj;
      });
    } else {
      // No subchecks in UI, preserve existing if any
      subchecks = existing?.subchecks || [];
      if (subchecks.length > 0) {
        console.log(`Row [${idx}] has no subchecks in UI, preserving existing:`, subchecks);
      }
    }

    const refObj = {
      id: existing?.id || id,
      cell: $r.find("input[data-type='cell']").val().trim(),
      keyword,
      sheet: $r.find("select[data-type='sheet']").val(),
      type: $r.find("select[data-type='refType']").val(),
      attackNameCell: $r.find("input[data-type='attackNameCell']").val()?.trim() || "",
      critRangeCell: $r.find("input[data-type='critRangeCell']").val()?.trim() || "",
      damageCell: $r.find("input[data-type='damageCell']").val()?.trim() || "",
      value: "",
      subchecks
    };
    console.log(`Reference [${idx}]:`, refObj);
    return refObj;
  });

  console.log("Final refs to save:", JSON.parse(JSON.stringify(refs)));

  const sheetId = sheet.actor.getFlag("sheexcel_updated", "sheetId");

  if (sheetId) {
    const updated = await batchFetchValues(sheetId, refs);
    console.log("Batch fetched values:", updated);
    await sheet.actor.setFlag("sheexcel_updated", "cellReferences", updated);
    updated.forEach((ref, idx) => {
      sheet.element
        .find(`.sheexcel-reference-row[data-index="${idx}"] .sheexcel-reference-value`)
        .text(ref.value);
    });
  } else {
    await sheet.actor.setFlag("sheexcel_updated", "cellReferences", refs);
  }

  sheet.render(false);
}