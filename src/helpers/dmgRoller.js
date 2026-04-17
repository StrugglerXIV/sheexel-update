// dmgRoller.js

import { promptDmgBonus } from "./situational.js";
import { wrapSheexcelChatFlavor } from "./chatStyling.js";

// ——————————————————————————————————————————
// DICE SCALE  d2 → d4 → d6 → d8 → d10 → d12
// ——————————————————————————————————————————
const DIE_SCALE = [2, 4, 6, 8, 10, 12];

export function shiftFormula(formula, direction) {
  // direction: +1 = scale up, -1 = scale down
  // Replaces every dN in the formula with the next/previous face count in DIE_SCALE.
  // If a die is already at the edge of the scale, it stays there.
  return formula.replace(/(\d*)d(\d+)/gi, (match, count, faces) => {
    const f = parseInt(faces);
    const idx = DIE_SCALE.indexOf(f);
    if (idx === -1) return match;
    const nextIdx = Math.min(Math.max(0, idx + direction), DIE_SCALE.length - 1);
    return `${count}d${DIE_SCALE[nextIdx]}`;
  });
}

// ——————————————————————————————————————————
// 1) PRE-ROLL PROMPT: scale + optional bonus
// ——————————————————————————————————————————
function promptDmgWithScale(keyword, initialFormula) {
  return new Promise(resolve => {
    let currentFormula = initialFormula;

    const dialog = new Dialog({
      title: `${keyword} — Damage`,
      content: `
        <div style="font-family:'Fontin','Georgia',serif;">
          <div style="margin-bottom:8px;font-size:0.85em;color:#c7a86d;">Base formula</div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
            <button type="button" id="sheexcel-scale-down"
              style="background:#1a0003;border:1.5px solid #8b3a1a;color:#e07050;border-radius:5px;
                     padding:3px 12px;font-size:1em;font-weight:700;cursor:pointer;">▼</button>
            <code id="sheexcel-formula-display"
              style="flex:1;text-align:center;background:#0d0a06;border:1px solid #3a2e1a;
                     border-radius:4px;padding:4px 8px;color:#ffd700;font-size:1em;
                     letter-spacing:0.5px;">${currentFormula}</code>
            <button type="button" id="sheexcel-scale-up"
              style="background:#061a03;border:1.5px solid #3a6a1a;color:#80d060;border-radius:5px;
                     padding:3px 12px;font-size:1em;font-weight:700;cursor:pointer;">▲</button>
          </div>
          <div style="margin-bottom:6px;font-size:0.85em;color:#c7a86d;">Situational bonus (optional)</div>
          <input type="text" id="sheexcel-dmg-bonus" value=""
            placeholder="e.g. +3, 1d4"
            style="width:100%;background:#1a140f;border:1px solid #bfa05a;color:#c7a86d;
                   padding:5px 9px;border-radius:4px;font-size:0.95em;box-sizing:border-box;"/>
          <div style="margin-top:10px;margin-bottom:6px;font-size:0.85em;color:#c7a86d;">Roll mode</div>
          <div style="display:flex;gap:16px;margin-bottom:10px;">
            <label style="cursor:pointer;color:#c7a86d;"><input type="radio" name="dmg-mode" value="adv"> ▲ Advantage</label>
            <label style="cursor:pointer;color:#c7a86d;"><input type="radio" name="dmg-mode" value="norm" checked> = Normal</label>
            <label style="cursor:pointer;color:#c7a86d;"><input type="radio" name="dmg-mode" value="dis"> ▼ Disadvantage</label>
          </div>
          <label style="cursor:pointer;color:#c7a86d;display:flex;align-items:center;gap:8px;">
            <input type="checkbox" id="sheexcel-dmg-crit"> Critical Hit?
          </label>
        </div>`,
      buttons: {
        roll: {
          label: "Roll",
          callback: html => {
            const bonus = html.find("#sheexcel-dmg-bonus").val().trim();
            const advMode = html.find("input[name='dmg-mode']:checked").val() || "norm";
            const isCrit = html.find("#sheexcel-dmg-crit").prop("checked");
            resolve({ formula: currentFormula, bonus: bonus || "", advMode, isCrit });
          }
        },
        cancel: {
          label: "Cancel",
          callback: () => resolve(null)
        }
      },
      default: "roll",
      render: html => {
        html.find("#sheexcel-scale-down").on("click", () => {
          const shifted = shiftFormula(currentFormula, -1);
          if (shifted !== currentFormula) {
            currentFormula = shifted;
            html.find("#sheexcel-formula-display").text(currentFormula);
          }
        });
        html.find("#sheexcel-scale-up").on("click", () => {
          const shifted = shiftFormula(currentFormula, +1);
          if (shifted !== currentFormula) {
            currentFormula = shifted;
            html.find("#sheexcel-formula-display").text(currentFormula);
          }
        });
      }
    });
    dialog.render(true);
  });
}

// ——————————————————————————————————————————
// 2) BUILD A CLEAN FORMULA STRING  
// ——————————————————————————————————————————
function buildFormula(baseValue, bonusRaw) {
  let bonusTerm = "";
  if ( bonusRaw ) {
    Roll.create(bonusRaw); // will throw if invalid
    bonusTerm = (/^[+\-]/.test(bonusRaw) ? "" : "+") + bonusRaw;
  }
  return `${baseValue >= 0 ? "+" : ""}${baseValue}${bonusTerm}`;
}

// ——————————————————————————————————————————
// 3) DOUBLE EVERY POSITIVE DIE FACE & POSITIVE “+” MODIFIER  
// ——————————————————————————————————————————
function applyCritical(roll) {
  roll.terms.forEach((term, i) => {
    // Determine sign for this term
    const prev = roll.terms[i - 1];
    const op = (prev?.constructor.name === "OperatorTerm" ? prev.operator : "+");

    // — Dice: only double faces if op is “+”
    if (term instanceof Die) {
      if (op === "+") {
        term.results.forEach(r => { if (r.result > 0) r.result *= 2; });
      }
    }

    // — Static numbers: only double +n
    else if (term instanceof NumericTerm && term.number > 0) {
      if (op === "+") {
        term.number *= 2;
      }
    }
  });
}

// ——————————————————————————————————————————
// 4) RECOMPUTE TOTAL RESPECTING “+” & “–”  
// ——————————————————————————————————————————
function computeTotal(roll) {
  return roll.terms.reduce((sum, term, i) => {
    const prev = roll.terms[i - 1];
    const op = (prev?.constructor.name === "OperatorTerm" ? prev.operator : "+");

    // — Dice: sum faces, apply sign
    if (term instanceof Die) {
      const subtotal = term.results.reduce((a, r) => a + r.result, 0);
      return sum + (op === "-" ? -subtotal : subtotal);
    }

    // — Static numbers: apply sign
    if (term instanceof NumericTerm) {
      return sum + (op === "-" ? -term.number : term.number);
    }
    return sum;
  }, 0);
}

function getDamageDisplayData(roll) {
  const keptDice = [];
  const droppedDice = [];
  let positiveFlat = 0;
  let negativeFlat = 0;

  roll.terms.forEach((term, i) => {
    const prev = roll.terms[i - 1];
    const op = (prev?.constructor.name === "OperatorTerm" ? prev.operator : "+");

    if (term instanceof Die) {
      term.results.forEach(r => {
        if (r.discarded) droppedDice.push(r.result);
        else keptDice.push(r.result);
      });
      return;
    }

    if (term instanceof NumericTerm) {
      const signed = (op === "-" ? -term.number : term.number);
      if (signed > 0) positiveFlat += signed;
      if (signed < 0) negativeFlat += signed;
    }
  });

  const adjustedDice = [...keptDice];
  let boostedDieIndex = -1;
  if (positiveFlat > 0 && adjustedDice.length) {
    const highest = Math.max(...adjustedDice);
    boostedDieIndex = adjustedDice.indexOf(highest);
    adjustedDice[boostedDieIndex] += positiveFlat;
  }

  const adjustedTotal = adjustedDice.reduce((sum, value) => sum + value, 0) + negativeFlat;

  return {
    keptDice,
    droppedDice,
    adjustedDice,
    boostedDieIndex,
    positiveFlat,
    negativeFlat,
    adjustedTotal
  };
}

function formatDiceList(values) {
  return `[${values.join(", ")}]`;
}

// ——————————————————————————————————————————
// 5) SEND THE CHAT MESSAGE  
// ——————————————————————————————————————————
async function postToChat(roll, preCrit, baseFormula, doubledFormula, isCrit, actor, damageType = "", rollMode = "norm") {
  const speaker = ChatMessage.getSpeaker({ actor });
  const damageLabel = damageType ? `Damage (${damageType})` : "Damage";
  const display = getDamageDisplayData(roll);

  const modeLabel = rollMode === "adv" ? "Advantage" : (rollMode === "dis" ? "Disadvantage" : "Normal");
  const rolledDiceLine = `<em>Rolled Dice:</em> ${formatDiceList(display.keptDice)}`;
  const afterBonusLine = `<br><em>After Bonuses:</em> ${formatDiceList(display.adjustedDice)}`;

  const droppedLine = display.droppedDice.length
    ? `<br><em>Dropped Dice:</em> ${formatDiceList(display.droppedDice)} (${modeLabel})`
    : "";

  const negativeLine = display.negativeFlat < 0
    ? `<br><em>Flat penalty:</em> ${display.negativeFlat}`
    : "";

  const flavor = isCrit
    ? `<strong>🔪 Critical ${damageLabel}</strong><br><em>Formula:</em> <code>${doubledFormula}</code><br>${rolledDiceLine}${afterBonusLine}${droppedLine}${negativeLine}`
    : `<strong>🔪 ${damageLabel}</strong><br><em>Formula:</em> <code>${baseFormula}</code><br>${rolledDiceLine}${afterBonusLine}${droppedLine}${negativeLine}`;

  await roll.toMessage({
    speaker,
    flavor: wrapSheexcelChatFlavor(flavor, "damage")
  });
}

// ——————————————————————————————————————————
// ADVANTAGE DAMAGE ROLLING (double dice, drop lowest half)
// ——————————————————————————————————————————
function doubleDiceInFormula(formula) {
  console.log("Input formula:", formula);
  // Add 'i' flag for case-insensitive match
  const regex = /(\b)(\d*)d(\d+)(?![a-zA-Z])/gi;
  let match;
  let matches = [];
  while ((match = regex.exec(formula)) !== null) {
    matches.push(match);
  }
  console.log("Regex matches:", matches);

  const result = formula.replace(regex, (full, pre, n, die) => {
    console.log("Match found:", { full, pre, n, die });
    const num = n === "" ? 1 : parseInt(n);
    const doubled = `${pre}${num * 2}d${die}`;
    console.log("Reeeplacing with:", doubled);
    return doubled;
  });

  console.log("Output formula:", result);
  return result;
}

function dropLowestDice(roll) {
  // For each Die term, drop half the lowest results (rounded down)
  roll.terms.forEach(term => {
    if (term instanceof Die) {
      const numToDrop = Math.floor(term.results.length / 2);
      // Sort results ascending, mark the lowest as dropped
      const sorted = [...term.results].sort((a, b) => a.result - b.result);
      for (let i = 0; i < numToDrop; i++) {
        sorted[i].discarded = true;
      }
    }
  });
}

function dropHighestDice(roll) {
  // For each Die term, drop half the highest results (rounded down)
  roll.terms.forEach(term => {
    if (term instanceof Die) {
      const numToDrop = Math.floor(term.results.length / 2);
      // Sort results descending, mark the highest as dropped
      const sorted = [...term.results].sort((a, b) => b.result - a.result);
      for (let i = 0; i < numToDrop; i++) {
        sorted[i].discarded = true;
      }
    }
  });
}

/**
 * Roll damage with advantage/disadvantage/normal
 * Usage: handleDamage({ ..., advantage: true }) or handleDamage({ ..., disadvantage: true })
 */
export async function handleDamage({ dmgF, isCrit, keyword, sheet, dmgAdvantage, dmgDisadvantage, damageType = "", bonusRawOverride = undefined }) {
  if (!dmgF) return false;
  // 6a) Bonus prompt (with pre-roll scale)
  let scaledDmgF = dmgF;
  let bonusRaw;
  let isCritLocal = isCrit;
  let isAdv = !!dmgAdvantage;
  let isDis = !!dmgDisadvantage;
  if (bonusRawOverride !== undefined) {
    bonusRaw = bonusRawOverride;
  } else {
    const result = await promptDmgWithScale(keyword, dmgF);
    if (result === null) return false; // user cancelled
    scaledDmgF = result.formula;
    bonusRaw = result.bonus;
    isCritLocal = result.isCrit;
    isAdv = result.advMode === "adv";
    isDis = result.advMode === "dis";
  }

  // 6b) Build & roll
  let baseFormula = buildFormula(scaledDmgF, bonusRaw);
  let formulaToRoll = baseFormula;
  if (isAdv || isDis) formulaToRoll = doubleDiceInFormula(baseFormula);
  console.log(formulaToRoll);

  const dmgRoll = new Roll(formulaToRoll);
  await dmgRoll.evaluate({ async: true });

  // 6c) Advantage/Disadvantage: drop half dice
  if (isAdv) dropLowestDice(dmgRoll);
  if (isDis) dropHighestDice(dmgRoll);

  // Recompute total to ignore discarded dice
  dmgRoll._total = dmgRoll.terms.reduce((sum, term, i) => {
    const prev = dmgRoll.terms[i - 1];
    const op = (prev?.constructor.name === "OperatorTerm" ? prev.operator : "+");

    if (term instanceof Die) {
      // Only sum non-discarded results
      const subtotal = term.results.filter(r => !r.discarded).reduce((a, r) => a + r.result, 0);
      return sum + (op === "-" ? -subtotal : subtotal);
    }
    if (term instanceof NumericTerm) {
      return sum + (op === "-" ? -term.number : term.number);
    }
    return sum;
  }, 0);

  // 6c) Capture pre-crit dice
  const preCrit = dmgRoll.terms
    .filter(t => t instanceof Die)
    .flatMap(t => t.results.map(r => r.result));

  let doubledFormula = baseFormula;
  // 6d) Crit handling
  if (isCritLocal) {
    applyCritical(dmgRoll);
    dmgRoll._total = computeTotal(dmgRoll);

    // Rebuild doubled formula showing actual modified dice
    doubledFormula = dmgRoll.terms.reduce((parts, term, i, terms) => {
      if (term.constructor.name === "OperatorTerm") return parts;
      const prev = terms[i - 1];
      const op = (prev?.constructor.name === "OperatorTerm" ? prev.operator : "+");
      if (term instanceof Die) {
        const faces = term.results.map(r => r.result);
        return [...parts, `${op}[${faces.join(",")}]`];
      }
      if (term instanceof NumericTerm) {
        return [...parts, `${op}${term.number}`];
      }
      return parts;
    }, []).join("");
  }

  // 6e) Post
  const rollMode = isAdv ? "adv" : (isDis ? "dis" : "norm");
  await postToChat(dmgRoll, preCrit, baseFormula, doubledFormula, isCritLocal, sheet.actor, damageType, rollMode);
  return true;
}
