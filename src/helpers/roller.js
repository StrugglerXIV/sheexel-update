// helpers/roller.js
import { handleDamage } from "./dmgRoller.js";
import { MODULE_NAME, FLAGS } from "./constants.js";
import { wrapSheexcelChatFlavor } from "./chatStyling.js";

// ——————————————————————————————————————————
// ATTACK CARD — Post card to chat, buttons roll later
// ——————————————————————————————————————————

function promptAttackOptions(keyword) {
  return new Promise(resolve => {
    new Dialog({
      title: `${keyword} — Attack`,
      content: `
        <div style="font-family:'Fontin','Georgia',serif;padding:4px 0;">
          <div style="margin-bottom:8px;font-size:0.85em;color:#c7a86d;">Roll mode</div>
          <div style="display:flex;gap:16px;margin-bottom:12px;">
            <label style="cursor:pointer;color:#c7a86d;"><input type="radio" name="atk-mode" value="adv"> ▲ Advantage</label>
            <label style="cursor:pointer;color:#c7a86d;"><input type="radio" name="atk-mode" value="norm" checked> = Normal</label>
            <label style="cursor:pointer;color:#c7a86d;"><input type="radio" name="atk-mode" value="dis"> ▼ Disadvantage</label>
          </div>
          <div style="margin-bottom:6px;font-size:0.85em;color:#c7a86d;">Situational bonus (optional)</div>
          <input type="text" id="sheexcel-atk-bonus" value="" placeholder="e.g. +3, 1d4"
            style="width:100%;background:#1a140f;border:1px solid #bfa05a;color:#c7a86d;
                   padding:5px 9px;border-radius:4px;font-size:0.95em;box-sizing:border-box;"/>
        </div>`,
      buttons: {
        roll: {
          label: "Roll",
          callback: html => {
            const advMode = html.find("input[name='atk-mode']:checked").val() || "norm";
            const bonus = html.find("#sheexcel-atk-bonus").val().trim();
            resolve({ advMode, bonus: bonus || "" });
          }
        },
        cancel: { label: "Cancel", callback: () => resolve(null) }
      },
      default: "roll"
    }).render(true);
  });
}

function promptCheckSaveOptions(keyword) {
  return new Promise(resolve => {
    new Dialog({
      title: `${keyword} — Roll`,
      content: `
        <div style="font-family:'Fontin','Georgia',serif;padding:4px 0;">
          <div style="margin-bottom:8px;font-size:0.85em;color:#c7a86d;">Roll mode</div>
          <div style="display:flex;gap:16px;margin-bottom:12px;">
            <label style="cursor:pointer;color:#c7a86d;"><input type="radio" name="roll-mode" value="adv"> ▲ Advantage</label>
            <label style="cursor:pointer;color:#c7a86d;"><input type="radio" name="roll-mode" value="norm" checked> = Normal</label>
            <label style="cursor:pointer;color:#c7a86d;"><input type="radio" name="roll-mode" value="dis"> ▼ Disadvantage</label>
          </div>
          <div style="margin-bottom:6px;font-size:0.85em;color:#c7a86d;">Situational bonus (optional)</div>
          <input type="text" id="sheexcel-roll-bonus" value="" placeholder="e.g. +3, 1d4"
            style="width:100%;background:#1a140f;border:1px solid #bfa05a;color:#c7a86d;
                   padding:5px 9px;border-radius:4px;font-size:0.95em;box-sizing:border-box;"/>
        </div>`,
      buttons: {
        roll: {
          label: "Roll",
          callback: html => {
            const advMode = html.find("input[name='roll-mode']:checked").val() || "norm";
            const bonus = html.find("#sheexcel-roll-bonus").val().trim();
            resolve({ advMode, bonus: bonus || "" });
          }
        },
        cancel: { label: "Cancel", callback: () => resolve(null) }
      },
      default: "roll"
    }).render(true);
  });
}

export async function handleAttackCard(event, sheet) {
  event.preventDefault();
  event.stopImmediatePropagation();
  const btn = $(event.currentTarget);
  const mod = Number(btn.data("value")) || 0;
  const crit = Number(btn.data("crit")) || 20;
  const dmgF = btn.data("damage") != null ? String(btn.data("damage")).trim() : "";
  const damageType = normalizeDamageType(btn.data("damageType"));
  const refIndex = btn.data("refIndex");
  const attackName = btn.closest(".attack-entry").find(".attack-name").text().trim();
  const accuracyText = btn.closest(".attack-entry").find(".attack-accuracy").text().trim();

  let damageParts = [];
  const refIdx = Number(refIndex);
  if (Number.isFinite(refIdx)) {
    const refs = sheet.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || [];
    const ref = refs[refIdx];
    if (Array.isArray(ref?.damageParts)) {
      damageParts = ref.damageParts
        .filter(part => part && String(part.formula || "").trim())
        .map(part => ({
          formula: String(part.formula || "").trim(),
          type: normalizeDamageType(part.type)
        }));
    }
  }

  const keyword = attackName || "Attack";
  const partsAttr = encodeURIComponent(JSON.stringify(damageParts));

  const dmgDisplay = damageParts.length > 0
    ? damageParts.map(p => `<span class="sheexcel-card-part"><code>${p.formula}</code>${p.type ? ` <em>${p.type}</em>` : ""}</span>`).join(" ")
    : (dmgF ? `<code>${dmgF}</code>${damageType ? ` <em>${damageType}</em>` : ""}` : "<em>No damage</em>");

  const content = `
    <div class="sheexcel-attack-card"
         data-mod="${mod}"
         data-crit="${crit}"
         data-dmg-f="${dmgF.replace(/"/g, "&quot;")}"
         data-damage-type="${damageType.replace(/"/g, "&quot;")}"
         data-damage-parts="${partsAttr}"
         data-actor-id="${sheet.actor.id}"
         data-ref-index="${refIdx}"
         data-keyword="${keyword.replace(/"/g, "&quot;")}">
      <div class="sheexcel-card-title">${keyword}</div>
      <div class="sheexcel-card-row">${accuracyText} &nbsp;·&nbsp; Crit ≥ ${crit}</div>
      <div class="sheexcel-card-row sheexcel-card-dmg-row">${dmgDisplay}</div>
      <div class="sheexcel-card-btns">
        <button type="button" class="sheexcel-attack-btn" data-action="attack">⚔ Attack</button>
        <button type="button" class="sheexcel-attack-btn" data-action="damage">🔪 Damage</button>
      </div>
    </div>`;

  await ChatMessage.create({
    speaker: ChatMessage.getSpeaker({ actor: sheet.actor }),
    content
  });
}

export async function rollAttackFromCard({ mod, crit, keyword, sheet }) {
  const opts = await promptAttackOptions(keyword);
  if (!opts) return;
  const { advMode, bonus: bonusRaw } = opts;

  let bonusTerm = "";
  if (bonusRaw) {
    try { Roll.create(bonusRaw); } catch { return ui.notifications.error("Invalid bonus: " + bonusRaw); }
    bonusTerm = (/^[+\-]/.test(bonusRaw) ? "" : "+") + bonusRaw;
  }

  let d20;
  if (advMode === "adv") d20 = "2d20kh1";
  else if (advMode === "dis") d20 = "2d20kl1";
  else d20 = "1d20";

  const formula = `${d20}${mod >= 0 ? "+" : ""}${mod}${bonusTerm}`;
  const roll = new Roll(formula);
  await roll.evaluate({ async: true });

  const terms = roll.terms.filter(t => t.faces === 20);
  const rolls = terms.flatMap(t => t.results.map(r => r.result));
  const isCrit = Math.max(...rolls) >= crit;

  await roll.toMessage({
    speaker: ChatMessage.getSpeaker({ actor: sheet.actor }),
    flavor: wrapSheexcelChatFlavor(
      `<strong>${keyword}</strong> → ${roll.total}` +
      (isCrit ? ` <span class="sheexcel-crit">[CRIT!]</span>` : ""),
      "attack"
    )
  });

  if (/^initiative$/i.test(keyword)) {
    await applyInitiativeToCombat(sheet, roll.total);
  }
}

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
  const rollKind = String(btn.data("rollKind") || "").trim().toLowerCase();
  let damageType = normalizeDamageType(btn.data("damageType"));
  const refIndex = Number(btn.data("refIndex"));
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
  } else if (rollKind === "save") {
    keyword = `${btn.text().trim()} Save`;
  } else {
    // For saves, attacks, spells: use button text
    keyword = btn.text().trim();
  }

  const totalMod = mod;

  const rollOptions = await promptCheckSaveOptions(keyword);
  if (!rollOptions) return;

  const { advMode, bonus: bonusRaw = "" } = rollOptions;

  // Validate by trying to build a Roll
  let bonusTerm = "";
  try {
    if (bonusRaw) {
      Roll.create(bonusRaw);
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

  const formula = `${d20}${totalMod >= 0 ? "+" : ""}${totalMod}${bonusTerm}`;
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
    flavor: wrapSheexcelChatFlavor(
      `<strong>${keyword}</strong> → ${roll.total}${attackName ? ` from ${attackName}` : ""}` +
      (isCrit ? ` <span class=\"sheexcel-crit\">[CRIT!]</span>` : ""),
      "check"
    )
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
        bonusRawOverride: undefined,
        sheet
      });

      if (result === false) {
        return;
      }
    }
  } else if (dmgF) {
    await handleDamage({
      dmgF,        // your base damage number
      isCrit,      // boolean from your attack-roll result
      keyword: damageType ? `Damage (${damageType})` : "Damage",
      damageType,
      sheet        // your ItemSheet instance (so we can get sheet.actor)
    });
  }
}
