import { apiCache } from './apiCache.js';
import { sanitizeSheetName, validateCellReference, ValidationError } from './validation.js';
import { ERROR_MESSAGES, REFERENCE_TYPES } from './constants.js';

export async function batchFetchValues(sheetId, refs) {
  if (!sheetId || !Array.isArray(refs)) {
    throw new ValidationError("Invalid parameters for batch fetch");
  }

  // Helper to recursively collect all queries with a path
  function collectQueries(refs, path = []) {
    let queries = [];
    refs.forEach((r, idx) => {
      try {
        const sheetName = sanitizeSheetName(r.sheet);

        // Main value
        if (r.cell) {
          const validCell = validateCellReference(r.cell);
          queries.push({ path: [...path, idx], field: "value", range: `${sheetName}!${validCell}` });
        }
        
        // Attack-specific fields
        if (r.type === REFERENCE_TYPES.ATTACKS) {
          if (r.attackNameCell) {
            const validCell = validateCellReference(r.attackNameCell);
            queries.push({ path: [...path, idx], field: "attackName", range: `${sheetName}!${validCell}` });
          }
          if (r.critRangeCell) {
            const validCell = validateCellReference(r.critRangeCell);
            queries.push({ path: [...path, idx], field: "critRange", range: `${sheetName}!${validCell}` });
          }
          if (r.damageCell) {
            const validCell = validateCellReference(r.damageCell);
            queries.push({ path: [...path, idx], field: "damage", range: `${sheetName}!${validCell}` });
          }
        }

        if (r.type === REFERENCE_TYPES.SPELLS) {
          const spellFields = [
            ["spellNameCell", "spellName"],
            ["circleCell", "circle"],
            ["spellTypeCell", "spellType"],
            ["componentsCell", "components"],
            ["castTimeCell", "castTime"],
            ["costCell", "cost"],
            ["rangeCell", "range"],
            ["durationCell", "duration"],
            ["descriptionCell", "description"],
            ["effectCell", "effect"],
            ["empowerCell", "empower"],
            ["sourceCell", "source"],
            ["disciplineCell", "discipline"]
          ];
          spellFields.forEach(([cellField, valueField]) => {
            if (!r[cellField]) return;
            const validCell = validateCellReference(r[cellField]);
            queries.push({ path: [...path, idx], field: valueField, range: `${sheetName}!${validCell}` });
          });
        }
        
        // Recurse into subchecks
        if (Array.isArray(r.subchecks) && r.subchecks.length) {
          queries = queries.concat(collectQueries(r.subchecks, [...path, idx, "subchecks"]));
        }
      } catch (error) {
        console.warn(`❌ Sheexcel | Invalid reference at index ${idx}:`, error.message);
        // Skip invalid references but continue processing others
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
    ref[path[path.length - 1]][field] = value || "";
  }

  // Collect all queries
  const queries = collectQueries(refs);

  if (!queries.length) {
    console.warn("❌ Sheexcel | No valid cell references found");
    return refs;
  }

  try {
    // Use cached API with automatic retry and fallback
    const ranges = queries.map(q => q.range);
    const json = await apiCache.batchGet(sheetId, ranges);

    // Assign the fetched values back to the correct fields in each ref/subcheck
    const extractValue = (vr) => {
      const values = Array.isArray(vr.values) ? vr.values : [];
      if (!values.length) return "";
      if (values.length === 1 && values[0].length === 1) return values[0][0] ?? "";
      const lines = values.map(row => row.filter(v => String(v ?? "").trim() !== "").join(" "))
        .filter(line => line.trim() !== "");
      return lines.join("\n");
    };

    if (json.valueRanges) {
      json.valueRanges.forEach((vr, i) => {
        const { path, field } = queries[i];
        let value = extractValue(vr);
        if (field === "components") {
          value = value.split("\n").filter(Boolean).join(", ");
        }
        assignByPath(refs, path, field, value);
      });
    }

    return refs;
  } catch (error) {
    console.error("❌ Sheexcel | Batch fetch failed:", error);
    
    // Set all values to empty string on failure
    queries.forEach(({ path, field }) => {
      assignByPath(refs, path, field, "");
    });
    
    throw new Error(`${ERROR_MESSAGES.API_ERROR}: ${error.message}`);
  }
}