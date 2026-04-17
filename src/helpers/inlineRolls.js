const INLINE_ROLL_REGEX = /\b\d+\s*[dD]\s*\d+(?:\s*[+\-]\s*\d+)?\b/g;
import { wrapSheexcelChatFlavor } from "./chatStyling.js";

const DIE_SCALE = [2, 4, 6, 8, 10, 12];

function normalizeFormula(formula) {
  return String(formula || "").replace(/\s+/g, "").replace(/D/g, "d");
}

function shiftFormula(formula, direction) {
  return String(formula || "").replace(/(\d*)d(\d+)/gi, (match, count, faces) => {
    const faceCount = Number.parseInt(faces, 10);
    const scaleIndex = DIE_SCALE.indexOf(faceCount);
    if (scaleIndex === -1) return match;

    const nextIndex = Math.min(Math.max(0, scaleIndex + direction), DIE_SCALE.length - 1);
    return `${count}d${DIE_SCALE[nextIndex]}`;
  });
}

function buildBonusTerm(rawBonus) {
  const bonus = String(rawBonus || "").trim();
  if (!bonus) return "";

  Roll.create(bonus); // Throws if invalid.
  return /^[+\-]/.test(bonus) ? bonus : `+${bonus}`;
}

function promptInlineRollOptions(contextLabel, initialFormula) {
  return new Promise((resolve) => {
    let currentFormula = normalizeFormula(initialFormula);

    const titleContext = String(contextLabel || "").trim() || "Description";
    const dialog = new Dialog({
      title: `${titleContext} — Roll`,
      content: `
        <div style="font-family:'Fontin','Georgia',serif;padding:4px 0;">
          <div style="margin-bottom:8px;font-size:0.85em;color:#c7a86d;">Base formula</div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
            <button type="button" id="sheexcel-inline-scale-down"
              style="background:#1a0003;border:1.5px solid #8b3a1a;color:#e07050;border-radius:5px;
                     padding:3px 12px;font-size:1em;font-weight:700;cursor:pointer;">▼</button>
            <code id="sheexcel-inline-formula-display"
              style="flex:1;text-align:center;background:#0d0a06;border:1px solid #3a2e1a;
                     border-radius:4px;padding:4px 8px;color:#ffd700;font-size:1em;
                     letter-spacing:0.5px;">${escapeHtml(currentFormula)}</code>
            <button type="button" id="sheexcel-inline-scale-up"
              style="background:#061a03;border:1.5px solid #3a6a1a;color:#80d060;border-radius:5px;
                     padding:3px 12px;font-size:1em;font-weight:700;cursor:pointer;">▲</button>
          </div>
          <div style="margin-bottom:6px;font-size:0.85em;color:#c7a86d;">Situational bonus (optional)</div>
          <input type="text" id="sheexcel-inline-roll-bonus" value="" placeholder="e.g. +3, 1d4"
            style="width:100%;background:#1a140f;border:1px solid #bfa05a;color:#c7a86d;
                   padding:5px 9px;border-radius:4px;font-size:0.95em;box-sizing:border-box;"/>
        </div>`,
      buttons: {
        roll: {
          label: "Roll",
          callback: (html) => {
            const bonus = html.find("#sheexcel-inline-roll-bonus").val().trim();
            resolve({ formula: currentFormula, bonus });
          }
        },
        cancel: {
          label: "Cancel",
          callback: () => resolve(null)
        }
      },
      default: "roll",
      render: (html) => {
        html.find("#sheexcel-inline-scale-down").on("click", () => {
          const shifted = shiftFormula(currentFormula, -1);
          if (shifted === currentFormula) return;

          currentFormula = shifted;
          html.find("#sheexcel-inline-formula-display").text(currentFormula);
        });

        html.find("#sheexcel-inline-scale-up").on("click", () => {
          const shifted = shiftFormula(currentFormula, +1);
          if (shifted === currentFormula) return;

          currentFormula = shifted;
          html.find("#sheexcel-inline-formula-display").text(currentFormula);
        });
      }
    });

    dialog.render(true);
  });
}

export function escapeHtml(value) {
  const source = String(value ?? "");

  if (typeof globalThis.Handlebars?.escapeExpression === "function") {
    return globalThis.Handlebars.escapeExpression(source);
  }

  return source
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildInlineRollSpan(match, actorId = "", contextLabel = "") {
  const formula = normalizeFormula(match);
  const actorAttr = actorId ? ` data-actor-id="${escapeHtml(String(actorId))}"` : "";
  const contextAttr = contextLabel ? ` data-context="${escapeHtml(String(contextLabel))}"` : "";
  return `<span class="sheexcel-inline-roll" data-formula="${escapeHtml(formula)}"${actorAttr}${contextAttr} role="button" tabindex="0">${escapeHtml(match)}</span>`;
}

export function enrichTextWithInlineRolls(text, { actorId = "", allowHtml = false, contextLabel = "" } = {}) {
  if (text == null) return "";

  const source = String(text);
  const prepared = allowHtml
    ? source
    : escapeHtml(source).replace(/\r?\n/g, "<br>");

  return prepared.replace(INLINE_ROLL_REGEX, (match) => buildInlineRollSpan(match, actorId, contextLabel));
}

export async function triggerInlineRoll(formula, actor = null, contextLabel = "") {
  const normalized = normalizeFormula(formula);
  if (!normalized) return;

  const opts = await promptInlineRollOptions(contextLabel, normalized);
  if (!opts) return;

  let finalFormula = opts.formula;
  try {
    finalFormula = `${opts.formula}${buildBonusTerm(opts.bonus)}`;
  } catch (_error) {
    ui.notifications?.error?.("Invalid bonus formula.");
    return;
  }

  const roll = new Roll(finalFormula);
  await roll.evaluate({ async: true });
  await roll.toMessage({
    speaker: ChatMessage.getSpeaker({ actor }),
    flavor: wrapSheexcelChatFlavor(
      contextLabel
        ? `<strong>Roll</strong> ${escapeHtml(finalFormula)} <span class="muted">from ${escapeHtml(contextLabel)}</span>`
        : `<strong>Roll</strong> ${escapeHtml(finalFormula)}`,
      "inline"
    )
  });
}

export async function handleInlineRollInteraction(event, actorResolver = null, contextResolver = null) {
  event.preventDefault();
  event.stopPropagation();

  if (event.type === "keydown" && !["Enter", " ", "Spacebar"].includes(event.key)) {
    return;
  }

  const target = event.target?.closest?.(".sheexcel-inline-roll") || event.currentTarget;
  const formula = target?.dataset?.formula;
  if (!formula) return;

  let actor = null;
  if (typeof actorResolver === "function") {
    actor = actorResolver(target) || null;
  }

  if (!actor) {
    const actorId = target?.dataset?.actorId;
    actor = actorId ? game.actors?.get(actorId) || null : null;
  }

  let contextLabel = String(target?.dataset?.context || "").trim();
  if (!contextLabel && typeof contextResolver === "function") {
    contextLabel = String(contextResolver(target) || "").trim();
  }

  await triggerInlineRoll(formula, actor, contextLabel);
}