// helpers/roller.js
import { handleDamage } from "./dmgRoller.js";
import { promptBonus } from "./situational.js";
import { MODULE_NAME, FLAGS } from "./constants.js";

function normalizeDamageType(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^(none|null|n\/?a|na|-)$/i.test(raw)) return "";
  return raw;
}

async function applyInitiativeToCombat(sheet, total) {
  const combat = game.combat;
  if (!combat) return;

  const combatants = combat.combatants
    .filter((combatant) => combatant?.actorId === sheet.actor.id && combatant.initiative == null);

  if (!combatants.length) return;

  const updates = combatants.map((combatant) => ({
    _id: combatant.id,
    initiative: total
  }));

  await combat.updateEmbeddedDocuments("Combatant", updates);
}

/**
 * Handle click on any .sheexcel-roll button in the main tab.
 * Expects the button to carry:
 *    data-value          → the attack modifier
 *    data-crit           → the numeric crit threshold
 *    data-damage         → the damage formula (e.g. "1d8+STR")
 *    data-damage-type    → optional damage type label (e.g. "Slashing")
 */
export async function handleRoll(event, sheet) {
  event.preventDefault();
  const btn = $(event.currentTarget);
  const mod = Number(btn.data("value")) || 0;
  const crit = Number(btn.data("crit")) || 20;
  const dmgF = btn.data("damage") != null ? String(btn.data("damage")).trim() : null;
  let damageType = normalizeDamageType(btn.data("damageType"));
  const refIndex = Number(btn.data("refIndex"));
  const advMode = btn.closest(".sheexcel-sidebar").find("input[name='roll-mode']:checked").val();
  const dmgAdvMode = btn.closest(".sheexcel-sidebar").find(".sheexcel-damage-mode").val();
  const attackName = $(btn).closest(".attack-entry").find(".attack-name").text();

  let damageParts = [];
  if (Number.isFinite(refIndex)) {
    const refs = sheet.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || [];
    const ref = refs[refIndex];
    if (Array.isArray(ref?.damageParts)) {
      damageParts = ref.damageParts
        .filter(part => part && String(part.formula || "").trim())
        .map(part => ({
          formula: String(part.formula || "").trim(),
          type: normalizeDamageType(part.type)
        }));
    }
  }

  if (!damageType) {
    const detailText = btn.closest(".attack-entry").find(".attack-detail").text();
    const stripped = String(detailText || "").replace(/[()]/g, "");
    const knownTypes = [
      "acid", "bludgeoning", "cold", "fire", "force", "lightning",
      "necrotic", "piercing", "poison", "psychic", "radiant", "slashing",
      "thunder", "vitality"
    ];

    const matches = knownTypes.filter((type) => new RegExp(`\\b${type}\\b`, "i").test(stripped));
    if (matches.length) {
      damageType = matches
        .map((type) => type.charAt(0).toUpperCase() + type.slice(1).toLowerCase())
        .join(" / ");
    }
  }

  // Use the button text as the keyword for the roll
  let keyword = "";
  const $currentEntry = btn.closest(".sheexcel-check-entry-sub, .sheexcel-check-entry");
  if ($currentEntry.length) {
    // For checks/subchecks
    keyword = $currentEntry.find(".sheexcel-check-keyword").first().text().trim();
  } else {
    // For saves, attacks, spells: use button text
    keyword = btn.text().trim();
  }

  const totalMod = mod;

  // Ask for the extra bonus formula
  const bonusRaw = await promptBonus(keyword);
  if (bonusRaw === null) return;
  if (!bonusRaw) bonusRaw = "";

  // Validate by trying to build a Roll
  let bonusTerm = "";
  try {
    if (bonusRaw) {
      // If they typed “3”, “+2”, “1d4+1”, “2d6”, etc.
      const testRoll = Roll.create(bonusRaw);
      // If that succeeded, we'll prefix a "+" if needed:
      bonusTerm = (bonusRaw.match(/^[+\-]/) ? "" : "+") + bonusRaw;
    }
  } catch (err) {
    return ui.notifications.error("Invalid bonus formula: " + bonusRaw);
  }

  // Build the d20 formula
  let d20;
  if (advMode === "adv") d20 = "2d20kh1";
  else if (advMode === "dis") d20 = "2d20kl1";
  else d20 = "1d20";

  const formula = `${d20}${totalMod >= 0 ? "+" : ""}${totalMod}${bonusTerm >= 0 ? "+" : ""}${bonusTerm}`;
  const roll = new Roll(formula);
  await roll.evaluate({ async: true });

  // Check for crit
  // extract the raw d20 result(s), before adding mod
  const terms = roll.terms.filter(t => t.faces === 20);
  const rolls = terms.flatMap(t => t.results.map(r => r.result));
  const top = Math.max(...rolls);
  const isCrit = top >= crit;

  // Render the d20 roll
  await roll.toMessage({
    speaker: ChatMessage.getSpeaker({ actor: sheet.actor }),
    flavor: `<strong>${keyword}</strong> → ${roll.total}${attackName ? ` from ${attackName}` : ""}` +
      (isCrit ? ` <span class=\"sheexcel-crit\">[CRIT!]</span>` : "")
  });

  if (/^initiative$/i.test(keyword)) {
    await applyInitiativeToCombat(sheet, roll.total);
  }

  // Roll damage if a formula is present
  if (damageParts.length > 1) {
    for (let i = 0; i < damageParts.length; i++) {
      const part = damageParts[i];
      const result = await handleDamage({
        dmgF: part.formula,
        isCrit,
        keyword: part.type ? `Damage (${part.type})` : `Damage Part ${i + 1}`,
        damageType: part.type,
        bonusRawOverride: i === 0 ? undefined : "",
        sheet,
        dmgAdvantage: dmgAdvMode === "advantage",
        dmgDisadvantage: dmgAdvMode === "disadvantage"
      });

      if (i === 0 && result === false) {
        return;
      }
    }
  } else if (dmgF) {
    await handleDamage({
      dmgF,        // your base damage number
      isCrit,      // boolean from your attack-roll result
      keyword: damageType ? `Damage (${damageType})` : "Damage",
      damageType,
      sheet,       // your ItemSheet instance (so we can get sheet.actor)
      dmgAdvantage: dmgAdvMode === "advantage",
      dmgDisadvantage: dmgAdvMode === "disadvantage"
    });
  }
}
