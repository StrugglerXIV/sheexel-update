// roller.js

import { promptDmgBonus } from "./situational.js";

// ——————————————————————————————————————————
// 1) PROMPT FOR AN OPTIONAL BONUS  
// ——————————————————————————————————————————
async function promptBonus(keyword) {
  const raw = await promptDmgBonus(keyword);
  if ( raw === null ) return null;               // user hit “cancel”
  const trimmed = raw.trim();
  return trimmed === "" ? "" : trimmed;       // normalize empty → ""
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

// ——————————————————————————————————————————
// 5) SEND THE CHAT MESSAGE  
// ——————————————————————————————————————————
async function postToChat(roll, preCrit, baseFormula, doubledFormula, isCrit, actor) {
  const speaker = ChatMessage.getSpeaker({ actor });
  const flavor = isCrit
    ? `<strong>🔪 Critical Damage</strong> (doubled): Rolled [${preCrit.join(", ")}] → Total ${roll.total}<br><em>Formula:</em> <code>${doubledFormula}</code>`
    : `<strong>🔪 Damage</strong>: ${baseFormula} = ${roll.total}`;
  await roll.toMessage({ speaker, flavor });
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
export async function handleDamage({ dmgF, isCrit, keyword, sheet, dmgAdvantage, dmgDisadvantage }) {
  if (!dmgF) return;
  // 6a) Bonus prompt
  const bonusRaw = await promptBonus(keyword);
  if (bonusRaw === null) return; // user cancelled

  // 6b) Build & roll
  let baseFormula = buildFormula(dmgF, bonusRaw);
  let formulaToRoll = baseFormula;
  let isAdv = !!dmgAdvantage;
  let isDis = !!dmgDisadvantage;
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
  if (isCrit) {
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
  await postToChat(dmgRoll, preCrit, baseFormula, doubledFormula, isCrit, sheet.actor);
}
