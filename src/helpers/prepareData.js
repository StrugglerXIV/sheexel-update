// helpers/prepareData.js
import { MODULE_NAME, FLAGS } from './constants.js';
import { enrichTextWithInlineRolls } from './inlineRolls.js';

function buildSheetOpenUrl(sheetUrl, sheetId) {
  const rawUrl = String(sheetUrl || '').trim();
  if (rawUrl) return rawUrl;
  if (!sheetId) return '';
  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/edit`;
}

function buildSheetEmbedUrl(sheetUrl, sheetId) {
  const fallback = sheetId
    ? `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/edit`
    : '';

  const rawUrl = String(sheetUrl || '').trim();
  if (!rawUrl) return fallback;

  try {
    // Use the URL as-is, stripping any previously stored minimal-UI params
    const url = new URL(rawUrl);
    url.searchParams.delete('rm');
    url.searchParams.delete('widget');
    url.searchParams.delete('headers');
    url.searchParams.delete('chrome');
    return url.toString();
  } catch (_error) {
    return fallback;
  }
}

/**
 * Enriches the Foundry-generated sheet data with Sheexcel flags and grouping.
 * @param {Object} data   The object returned by super.getData()
 * @param {Actor}  actor  The Actor document
 * @returns {Object}      The same data object, mutated with extra fields
 */
export async function prepareSheetData(data, actor) {
  const actorId = actor?.id || "";
  const abilityMods = actor.getFlag(MODULE_NAME, FLAGS.ABILITY_MODS) || {};
  const formatAbilityTokenText = (value) => {
    const text = String(value || "");
    if (!text) return "";

    return text.replace(/\b(STR|DEX|CON|INT|WIS|CHA)\b/g, (match) => {
      const mod = abilityMods[match];
      if (!Number.isFinite(mod)) return match;
      return `${match} [${mod}]`;
    });
  };
  const normalizeDamageType = (value) => {
    const raw = String(value || "").trim();
    if (!raw) return "";
    if (/^(none|null|n\/?a|na|-)$/i.test(raw)) return "";
    return raw;
  };

  const normalizeDamageParts = (parts) => {
    if (!Array.isArray(parts)) return [];

    const seen = new Set();
    return parts
      .map((part) => ({
        ...part,
        formula: String(part?.formula || "").trim(),
        type: normalizeDamageType(part?.type)
      }))
      .filter((part) => part.formula)
      .filter((part) => {
        const key = `${part.formula}::${String(part.type || "").toLowerCase()}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  };

  // Pull in your saved flags from the actor
  data.sheetUrl         = actor.getFlag(MODULE_NAME, FLAGS.SHEET_URL) || "";
  data.sheetId          = actor.getFlag(MODULE_NAME, FLAGS.SHEET_ID) || "";
  data.hideMenu         = actor.getFlag(MODULE_NAME, FLAGS.HIDE_MENU) || false;
  data.zoomLevel        = actor.getFlag(MODULE_NAME, FLAGS.ZOOM_LEVEL) ?? 100;
  data.sidebarCollapsed = actor.getFlag(MODULE_NAME, FLAGS.SIDEBAR_COLLAPSED) || false;
  const rawGearCurrency = actor.getFlag(MODULE_NAME, FLAGS.GEAR_CURRENCY) || {};
  data.gearCurrency = {
    onPerson: {
      gold: String(rawGearCurrency?.onPerson?.gold || "").trim(),
      silver: String(rawGearCurrency?.onPerson?.silver || "").trim(),
      copper: String(rawGearCurrency?.onPerson?.copper || "").trim()
    },
    banked: {
      gold: String(rawGearCurrency?.banked?.gold || "").trim(),
      silver: String(rawGearCurrency?.banked?.silver || "").trim(),
      copper: String(rawGearCurrency?.banked?.copper || "").trim()
    }
  };
  const rawRestEntries   = actor.getFlag(MODULE_NAME, FLAGS.REST_ENTRIES) || [];
  data.sheetOpenUrl     = buildSheetOpenUrl(data.sheetUrl, data.sheetId);
  data.sheetEmbedUrl    = buildSheetEmbedUrl(data.sheetUrl, data.sheetId);

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
    
    const damageParts = normalizeDamageParts(r.damageParts);
    const gearFlavorDisplay = formatAbilityTokenText(r.flavor || "");
    const gearDescriptionDisplay = formatAbilityTokenText(r.description || "");
    const gearAbilitiesDisplay = formatAbilityTokenText(r.abilities || "");
    const gearTypeDisplay = formatAbilityTokenText(r.gearType || "");
    const valueDisplay = formatAbilityTokenText(r.value || "");
    const weightDisplay = formatAbilityTokenText(r.weight || "");
    const bulkDisplay = formatAbilityTokenText(r.bulk || "");
    const powerDisplay = formatAbilityTokenText(r.power || "");
    const noiseDisplay = formatAbilityTokenText(r.noise || "");
    const durabilityDisplay = formatAbilityTokenText(r.durability || "");
    const integrityDisplay = formatAbilityTokenText(r.integrity || "");
    const resilienceDisplay = formatAbilityTokenText(r.resilience || "");
    const reachDisplay = formatAbilityTokenText(r.reach || "");
    const drawDisplay = formatAbilityTokenText(r.draw || "");
    const accuracyDisplay = formatAbilityTokenText(r.accuracy || "");
    const criticalDisplay = formatAbilityTokenText(r.critical || "");
    const damageDisplay = formatAbilityTokenText(r.damage || "");
    const armorDisplay = formatAbilityTokenText(r.armor || "");
    const donningDisplay = formatAbilityTokenText(r.donning || "");
    const discomfortDisplay = formatAbilityTokenText(r.discomfort || "");
    const chargesDisplay = formatAbilityTokenText(r.charges || "");
    const fuelDisplay = formatAbilityTokenText(r.fuel || "");
    const volumeDisplay = formatAbilityTokenText(r.volume || "");
    const singleHandDisplay = formatAbilityTokenText(r.singleHand || "");

    return {
      ...r,
      index: i,
      sheetNames,
      descriptionHtml: enrichTextWithInlineRolls(r.description || "", { actorId, contextLabel: r.abilityName || r.spellName || r.gearName || r.keyword || "Entry" }),
      effectHtml: enrichTextWithInlineRolls(r.effect || "", { actorId, contextLabel: r.abilityName || r.spellName || r.gearName || r.keyword || "Entry" }),
      notesHtml: enrichTextWithInlineRolls(r.notes || "", { actorId, contextLabel: r.abilityName || r.keyword || "Ability" }),
      empowerHtml: enrichTextWithInlineRolls(r.empower || "", { actorId, contextLabel: r.spellName || r.keyword || "Spell" }),
      gearFlavorDisplay,
      gearDescriptionDisplay,
      gearAbilitiesDisplay,
      gearTypeDisplay,
      valueDisplay,
      weightDisplay,
      bulkDisplay,
      powerDisplay,
      noiseDisplay,
      durabilityDisplay,
      integrityDisplay,
      resilienceDisplay,
      reachDisplay,
      drawDisplay,
      accuracyDisplay,
      criticalDisplay,
      damageDisplay,
      armorDisplay,
      donningDisplay,
      discomfortDisplay,
      chargesDisplay,
      fuelDisplay,
      volumeDisplay,
      singleHandDisplay,
      gearFlavorHtml: enrichTextWithInlineRolls(gearFlavorDisplay, { actorId, contextLabel: r.gearName || r.keyword || "Gear" }),
      gearDescriptionHtml: enrichTextWithInlineRolls(gearDescriptionDisplay, { actorId, contextLabel: r.gearName || r.keyword || "Gear" }),
      gearAbilitiesHtml: enrichTextWithInlineRolls(gearAbilitiesDisplay, { actorId, allowHtml: true, contextLabel: r.gearName || r.keyword || "Gear" }),
      damageType: normalizeDamageType(r.damageType),
      damageParts,
      hasMultiDamageParts: damageParts.length > 1,
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
    abilities: [],
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

  const abilities = data.groupedReferences.abilities || [];
  data.groupedReferences.abilities = abilities.sort((a, b) => {
    const aName = String(a?.abilityName || a?.keyword || "").trim();
    const bName = String(b?.abilityName || b?.keyword || "").trim();
    return aName.localeCompare(bName, undefined, { sensitivity: "base" });
  });

  const normalizeAbilityCategory = (value) => String(value || "").trim();
  const abilityGroupMap = new Map();
  data.groupedReferences.abilities.forEach((ability) => {
    const category = normalizeAbilityCategory(ability.category) || "Other";
    if (!abilityGroupMap.has(category)) abilityGroupMap.set(category, []);
    abilityGroupMap.get(category).push(ability);
  });

  data.groupedReferences.abilitiesByCategory = Array.from(abilityGroupMap.entries())
    .map(([label, list]) => ({
      label,
      abilities: list.sort((a, b) => {
        const aName = String(a?.abilityName || a?.keyword || "").trim();
        const bName = String(b?.abilityName || b?.keyword || "").trim();
        return aName.localeCompare(bName, undefined, { sensitivity: "base" });
      })
    }))
    .sort((a, b) => a.label.localeCompare(b.label, undefined, { sensitivity: "base" }));

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

  const restEntries = Array.isArray(rawRestEntries) ? rawRestEntries : [];
  const restSectionMap = new Map();

  restEntries.forEach((entry, index) => {
    const section = String(entry?.section || "Rest").trim() || "Rest";
    const title = String(entry?.title || "").trim();
    const postFullCard = /^(start of rest|end of rest)$/i.test(title);
    const normalized = {
      ...entry,
      index,
      section,
      title,
      postFullCard,
      summary: String(entry?.summary || "").trim(),
      summaryHtml: enrichTextWithInlineRolls(entry?.summary || "", { actorId, contextLabel: entry?.title || "Rest" }),
      details: Array.isArray(entry?.details)
        ? entry.details
          .map((detail) => {
            if (typeof detail === "string") {
              return { text: detail.trim(), html: enrichTextWithInlineRolls(detail, { actorId, contextLabel: entry?.title || "Rest" }), level: 1 };
            }

            return {
              text: String(detail?.text || "").trim(),
              html: enrichTextWithInlineRolls(detail?.text || "", { actorId, contextLabel: entry?.title || "Rest" }),
              level: Math.max(1, Number(detail?.level) || 1)
            };
          })
          .filter((detail) => detail.text)
        : []
    };

    if (!restSectionMap.has(section)) restSectionMap.set(section, []);
    restSectionMap.get(section).push(normalized);
  });

  data.restSheet = {
    entries: restEntries,
    sections: Array.from(restSectionMap.entries()).map(([label, entries]) => ({
      label,
      entries
    }))
  };

  return data;
}
