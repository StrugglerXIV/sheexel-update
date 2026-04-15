const INLINE_ROLL_REGEX = /\b\d+\s*[dD]\s*\d+(?:\s*[+\-]\s*\d+)?\b/g;

function normalizeFormula(formula) {
  return String(formula || "").replace(/\s+/g, "").replace(/D/g, "d");
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

  const roll = new Roll(normalized);
  await roll.evaluate({ async: true });
  await roll.toMessage({
    speaker: ChatMessage.getSpeaker({ actor }),
    flavor: contextLabel
      ? `<strong>Roll</strong> ${escapeHtml(normalized)} <span class="muted">from ${escapeHtml(contextLabel)}</span>`
      : `<strong>Roll</strong> ${escapeHtml(normalized)}`
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