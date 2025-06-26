// helpers/roller.js

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
  const dmgF   = String(btn.data("damage")) || "";
  const advMode= btn.closest(".sheexcel-sidebar").find("input[name='roll-mode']:checked").val();

  // build the d20 formula
  let d20;
  if (advMode === "adv") d20 = "2d20kh1";
  else if (advMode === "dis") d20 = "2d20kl1";
  else d20 = "1d20";

  const formula = `${d20}+${mod >= 0 ? mod : `(${mod})`}`;
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

  // If this was an attack AND a crit or not, immediately ask to roll damage
  if (dmgF) {
    const dmgRoll = new Roll(dmgF);
    await dmgRoll.evaluate({ async: true });
    await dmgRoll.toMessage({
      speaker: ChatMessage.getSpeaker({ actor: sheet.actor }),
      flavor: isCrit
        ? `<strong>🔪 Critical Damage</strong>: ${dmgRoll.formula}`
        : `<strong>🔪 Damage</strong>: ${dmgRoll.formula}`
    });
  }
}
