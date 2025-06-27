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
// 6) MAIN ENTRY POINT  
// ——————————————————————————————————————————
export async function handleDamage({ dmgF, isCrit, keyword, sheet }) {
  if (!dmgF) return;

  // 6a) Bonus prompt
  const bonusRaw = await promptBonus(keyword);
  if (bonusRaw === null) return; // user cancelled

  // 6b) Build & roll
  const baseFormula = buildFormula(dmgF, bonusRaw);
  const dmgRoll = new Roll(baseFormula);
  await dmgRoll.evaluate({ async: true });

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
