// helpers/prepareData.js
import { MODULE_NAME, FLAGS } from './constants.js';

/**
 * Enriches the Foundry-generated sheet data with Sheexcel flags and grouping.
 * @param {Object} data   The object returned by super.getData()
 * @param {Actor}  actor  The Actor document
 * @returns {Object}      The same data object, mutated with extra fields
 */
export async function prepareSheetData(data, actor) {
  // Pull in your saved flags from the actor
  data.sheetUrl         = actor.getFlag(MODULE_NAME, FLAGS.SHEET_URL) || "";
  data.hideMenu         = actor.getFlag(MODULE_NAME, FLAGS.HIDE_MENU) || false;
  data.zoomLevel        = actor.getFlag(MODULE_NAME, FLAGS.ZOOM_LEVEL) ?? 100;
  data.sidebarCollapsed = actor.getFlag(MODULE_NAME, FLAGS.SIDEBAR_COLLAPSED) || false;
  data.gearCurrency      = actor.getFlag(MODULE_NAME, FLAGS.GEAR_CURRENCY) || null;

  // Read refs & sheetNames from the actor
  const refs       = actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || [];
  const sheetNames = actor.getFlag(MODULE_NAME, FLAGS.SHEET_NAMES) || []; 

  // Build adjustedReferences with validation
  data.adjustedReferences = refs.map((r, i) => {
    // Ensure subchecks is always an array and sort alphabetically by keyword
    const subchecks = Array.isArray(r.subchecks) ? r.subchecks : [];
    const sortedSubchecks = subchecks
      .slice() // Create a copy to avoid mutating original
      .sort((a, b) => (a.keyword || '').localeCompare(b.keyword || '', undefined, { sensitivity: 'base' }));
    
    return {
      ...r,
      index: i,
      sheetNames,
      // Attack-specific fields with safe defaults
      attackNameCell: r.attackNameCell || "",
      critRangeCell:  r.critRangeCell  || "",
      damageCell:     r.damageCell     || "",
      // Ensure subchecks is properly structured and sorted
      subchecks: sortedSubchecks.map((sub, subIdx) => ({
        ...sub,
        subIndex: subIdx,
        parentIndex: i
      }))
    };
  });

  // Group into checks/saves/attacks/spells/gears with safe fallbacks
  data.groupedReferences = { 
    checks: [], 
    saves: [], 
    attacks: [], 
    spells: [],
    gears: []
  };
  
  data.adjustedReferences.forEach(r => {
    const targetArray = data.groupedReferences[r.type];
    if (targetArray) {
      targetArray.push(r);
    } else {
      // Fallback to checks for unknown types
      console.warn(`❌ Sheexcel | Unknown reference type: ${r.type}`);
      data.groupedReferences.checks.push(r);
    }
  });

  // Sort checks alphabetically, but always keep Initiative first
  data.groupedReferences.checks = (data.groupedReferences.checks || []).sort((a, b) => {
    const aKey = String(a?.keyword || "").trim();
    const bKey = String(b?.keyword || "").trim();
    const aInit = /^initiative$/i.test(aKey);
    const bInit = /^initiative$/i.test(bKey);
    if (aInit && !bInit) return -1;
    if (!aInit && bInit) return 1;
    return aKey.localeCompare(bKey, undefined, { sensitivity: "base" });
  });

  // Sort saves by canonical ability order
  const saveOrder = ["strength", "dexterity", "constitution", "intelligence", "wisdom", "charisma"];
  const saveOrderMap = new Map(saveOrder.map((name, index) => [name, index]));
  data.groupedReferences.saves = (data.groupedReferences.saves || []).sort((a, b) => {
    const aKey = String(a?.keyword || "").trim().toLowerCase();
    const bKey = String(b?.keyword || "").trim().toLowerCase();
    const aRank = saveOrderMap.has(aKey) ? saveOrderMap.get(aKey) : Number.MAX_SAFE_INTEGER;
    const bRank = saveOrderMap.has(bKey) ? saveOrderMap.get(bKey) : Number.MAX_SAFE_INTEGER;
    if (aRank !== bRank) return aRank - bRank;
    return aKey.localeCompare(bKey, undefined, { sensitivity: "base" });
  });

  // Group spells by circle for display
  const spells = data.groupedReferences.spells || [];
  const normalizeCircle = (value) => String(value ?? "").trim();
  const getCircleOrder = (value) => {
    const num = Number.parseInt(value, 10);
    return Number.isFinite(num) ? num : 9999;
  };

  const circleMap = new Map();
  spells.forEach(spell => {
    const circle = normalizeCircle(spell.circle);
    const label = circle ? `Circle ${circle}` : "Other";
    if (!circleMap.has(label)) circleMap.set(label, []);
    circleMap.get(label).push(spell);
  });

  const spellsByCircle = Array.from(circleMap.entries())
    .map(([label, list]) => ({
      label,
      order: getCircleOrder(label.replace(/^Circle\s+/i, "")),
      spells: list.sort((a, b) => (a.spellName || a.keyword || "").localeCompare(b.spellName || b.keyword || "", undefined, { sensitivity: "base" }))
    }))
    .sort((a, b) => a.order - b.order || a.label.localeCompare(b.label));

  data.groupedReferences.spellsByCircle = spellsByCircle;

  // Group gears by predefined categories based on type keywords
  const gears = data.groupedReferences.gears || [];
  
  const categoryDefs = [
    { key: 'weapons', label: 'Weapon', keywords: ['weapon'] },
    { key: 'armor', label: 'Armor', keywords: ['armor'] },
    { key: 'containers', label: 'Container', keywords: ['container'] },
    { key: 'tools', label: 'Tool', keywords: ['tool'] },
    { key: 'food', label: 'Food', keywords: ['food'] },
    { key: 'clothing', label: 'Clothing', keywords: ['clothing'] },
    { key: 'medicine', label: 'Medicine', keywords: ['medicine'] },
    { key: 'magical', label: 'Magical', keywords: ['magical'] },
    { key: 'jewelry', label: 'Jewelry', keywords: ['jewelry'] },
    { key: 'ammunition', label: 'Ammunition', keywords: ['ammunition'] },
    { key: 'instruments', label: 'Instrument', keywords: ['instrument'] }
  ];

  const gearsByCategory = {};
  categoryDefs.forEach(cat => {
    gearsByCategory[cat.key] = [];
  });
  gearsByCategory.uncategorized = [];

  gears.forEach(gear => {
    const typeStr = (gear.gearType || '').toLowerCase();
    let placed = false;
    
    for (const cat of categoryDefs) {
      if (cat.keywords.some(kw => typeStr.includes(kw.toLowerCase()))) {
        gearsByCategory[cat.key].push(gear);
        placed = true;
        break;
      }
    }
    
    if (!placed) {
      gearsByCategory.uncategorized.push(gear);
    }
  });

  const gearsByType = categoryDefs
    .filter(cat => gearsByCategory[cat.key].length > 0)
    .map(cat => ({
      label: cat.label,
      gears: gearsByCategory[cat.key].sort((a, b) => (a.gearName || a.keyword || "").localeCompare(b.gearName || b.keyword || "", undefined, { sensitivity: "base" }))
    }));

  if (gearsByCategory.uncategorized.length > 0) {
    gearsByType.push({
      label: 'Uncategorized',
      gears: gearsByCategory.uncategorized.sort((a, b) => (a.gearName || a.keyword || "").localeCompare(b.gearName || b.keyword || "", undefined, { sensitivity: "base" }))
    });
  }

  data.groupedReferences.gearsByType = gearsByType;

  return data;
}
