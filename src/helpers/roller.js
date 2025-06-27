// helpers/roller.js
import { promptBonus } from "./situational.js";
import { promptDmgBonus } from "./situational.js";
import { handleDamage } from "./dmgRoller.js";
/**
 * Handle click on any .sheexcel-roll button in the main tab.
 * Expects the button to carry:
 *    data-value          → the attack modifier
 *    data-crit           → the numeric crit threshold
 *    data-damage         → the damage formula (e.g. "1d8+STR")
 *    data-advantage-mode → "adv", "norm", or "dis"
 */
export async function handleRoll(event, sheet) {
  event.preventDefault();
  const btn    = $(event.currentTarget);
  const mod    = Number(btn.data("value")) || 0;
  const crit   = Number(btn.data("crit"))  || 20;
  const dmgF = btn.data("damage") != null ? String(btn.data("damage")).trim() : null;
  const advMode= btn.closest(".sheexcel-sidebar").find("input[name='roll-mode']:checked").val();

  // situational bonus inject
  const keyword = btn.text().trim();
  const totalMod = mod;
  
    // Ask for the extra bonus formula
  const bonusRaw = await promptBonus(keyword);
  if (bonusRaw === null) return;               // user cancelled entirely
  if (!bonusRaw) bonusRaw = "";                // empty string → no extra bonus
  
  // Validate by trying to build a Roll
  let bonusTerm = "";
  try {
    if (bonusRaw) {
      // If they typed “3”, “+2”, “1d4+1”, “2d6”, etc.
      const testRoll = Roll.create(bonusRaw);
      // If that succeeded, we'll prefix a "+" if needed:
      bonusTerm = (bonusRaw.match(/^[+\-]/) ? "" : "+") + bonusRaw;
    }
  } catch(err) {
    return ui.notifications.error("Invalid bonus formula: " + bonusRaw);
  }

  // build the d20 formula
  let d20;
  if (advMode === "adv") d20 = "2d20kh1";
  else if (advMode === "dis") d20 = "2d20kl1";
  else d20 = "1d20";

  const formula = `${d20}${totalMod>=0?"+":""}${totalMod}${bonusTerm>=0?"+":""}${bonusTerm}`;
  const roll    = new Roll(formula);
  await roll.evaluate({ async: true });

  // Check for crit
  // extract the raw d20 result(s), before adding mod
  const terms = roll.terms.filter(t => t.faces === 20);
  const rolls = terms.flatMap(t => t.results.map(r => r.result));
  const top   = Math.max(...rolls);
  const isCrit= top >= crit;

  // Render the d20 roll
  await roll.toMessage({
    speaker: ChatMessage.getSpeaker({ actor: sheet.actor }),
    flavor: `<strong>${btn.text()}</strong> → ${roll.total}` +
            (isCrit ? ` <span class="sheexcel-crit">[CRIT!]</span>` : "")
  });

	// wherever you had your big inline if (dmgF) { … }
	await handleDamage({
		dmgF,        // your base damage number
		isCrit,      // boolean from your attack-roll result
		keyword: "Damage",
		sheet        // your ItemSheet instance (so we can get sheet.actor)
});
}
