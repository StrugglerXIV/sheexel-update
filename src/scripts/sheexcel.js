import { MODULE_NAME, FLAGS, SETTINGS } from "../helpers/constants.js";
import { handleInlineRollInteraction } from "../helpers/inlineRolls.js";
import { rollAttackFromCard } from "../helpers/roller.js";
import { handleDamage } from "../helpers/dmgRoller.js";

// 1) Preload our Handlebars partials before anything renders
Hooks.once("init", async () => {
  const paths = [
    "modules/sheexcel_updated/templates/partials/main-tab.hbs",
    "modules/sheexcel_updated/templates/partials/sheet-tab.hbs",
    "modules/sheexcel_updated/templates/partials/rest-tab.hbs",
    "modules/sheexcel_updated/templates/partials/references-tab.hbs",
    "modules/sheexcel_updated/templates/partials/settings-tab.hbs"
  ];
  try {
    await loadTemplates(paths);
  } catch (err) {
    console.error("❌ Sheexcel | loadTemplates failed:", err, { paths });
    return; // Stop initialization if templates fail
  }
  
  // Google API Key setting
  game.settings.register(MODULE_NAME, SETTINGS.GOOGLE_API_KEY, {
    name: "Google Sheets API Key",
    hint: "Enter your Google Sheets API key here.",
    scope: "world",
    config: true,
    type: String,
    default: "",
  });

  // Damage modes setting (client)
  game.settings.register(MODULE_NAME, SETTINGS.DAMAGE_MODES, {
    name: "Attack Damage Modes",
    scope: "client", 
    config: false,
    type: Object,
    default: {}
  });
});

// --- Actor Sheet Class ---
import { SheexcelActorSheet } from "./SheexcelActorSheet.js";

// --- Sheet Registration ---
Hooks.once("setup", () => {
  Actors.registerSheet(MODULE_NAME, SheexcelActorSheet, {
    types: ["character", "npc", "creature", "vehicle"],
    label: "Sheexcel",
    makeDefault: false
  });
});

// --- Derived Data Wrapper ---
Hooks.once("ready", () => {
  if (typeof libWrapper !== "undefined") {
    libWrapper.register(
      MODULE_NAME,
      "CONFIG.Actor.documentClass.prototype.prepareDerivedData",
      function(wrapped, ...args) {
        wrapped(...args);
        const refs = this.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || [];
        this.system = this.system || {};
        this.system.sheexcel = refs.reduce((o, r) => {
          if (r.keyword) o[r.keyword] = r.value;
          return o;
        }, {});
      },
      "WRAPPER"
    );
  } else {
    // libWrapper not installed — patch natively as a fallback
    const _origPrepare = CONFIG.Actor.documentClass.prototype.prepareDerivedData;
    CONFIG.Actor.documentClass.prototype.prepareDerivedData = function(...args) {
      _origPrepare.apply(this, args);
      const refs = this.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || [];
      this.system = this.system || {};
      this.system.sheexcel = refs.reduce((o, r) => {
        if (r.keyword) o[r.keyword] = r.value;
        return o;
      }, {});
    };
  }

  $(document).on("click keydown", ".chat-message .sheexcel-inline-roll", (event) => {
    Promise.resolve(handleInlineRollInteraction(
      event,
      null,
      (target) => {
        const message = target.closest(".chat-message");
        const spellTitle = message?.querySelector(".sheexcel-spell-chat-title")?.textContent?.trim();
        if (spellTitle) return spellTitle;

        const restTitle = message?.querySelector(".sheexcel-spell-chat-sub")?.textContent?.trim();
        if (restTitle) return restTitle;

        return "Chat Card";
      }
    ))
      .catch((error) => {
        console.error("❌ Sheexcel | chat inline roll handler failed", error);
      });
  });
});

Hooks.on("renderChatMessage", (message, html) => {
  if (html.find(".sheexcel-chat-flavor").length) {
    html.addClass("sheexcel-chat-message sheexcel-chat-roll-message");
  }

  if (html.find(".sheexcel-chat-flavor-damage").length) {
    html.addClass("sheexcel-chat-damage-message");
    html.find(".dice-roll, .dice-tooltip").remove();

    const messageContent = html.find(".message-content");
    if (messageContent.length) {
      const isEmpty = messageContent.children().length === 0 && !messageContent.text().trim();
      if (isEmpty) messageContent.remove();
    }
  }

  if (html.find(".sheexcel-attack-card, .sheexcel-spell-chat").length) {
    html.addClass("sheexcel-chat-message sheexcel-chat-card-message");
  }

  const inlineRollCount = html.find(".sheexcel-inline-roll").length;
  if (!inlineRollCount) return;

  html.find(".sheexcel-inline-roll").each((_, el) => {
    const node = el;
    const handler = (ev) => {
      handleInlineRollInteraction(ev, null, (target) => {
        const spellTitle = html.find(".sheexcel-spell-chat-title").text().trim();
        if (spellTitle) return spellTitle;
        const restTitle = html.find(".sheexcel-spell-chat-sub").text().trim();
        if (restTitle) return restTitle;
        return "Chat Card";
      });
    };

    node.addEventListener("click", handler);
    node.addEventListener("keydown", handler);
  });
});

// ——————————————————————————————————————————
// ATTACK CARD — handle Attack / Damage buttons in chat
// ——————————————————————————————————————————
Hooks.on("renderChatMessage", (message, html) => {
  html.find(".sheexcel-attack-btn").on("click", async function () {
    const action = this.dataset.action;
    const card = $(this).closest(".sheexcel-attack-card")[0];
    if (!card) return;

    const actorId = card.dataset.actorId;
    const actor = game.actors?.get(actorId);
    if (!actor) return ui.notifications.warn("Actor not found for this attack card.");
    const fakeSheet = { actor };

    if (action === "attack") {
      const mod = parseInt(card.dataset.mod) || 0;
      const crit = parseInt(card.dataset.crit) || 20;
      const keyword = card.dataset.keyword || "Attack";
      await rollAttackFromCard({ mod, crit, keyword, sheet: fakeSheet });

    } else if (action === "damage") {
      const dmgF = card.dataset.dmgF || null;
      const damageType = card.dataset.damageType || "";
      let damageParts = [];
      try { damageParts = JSON.parse(decodeURIComponent(card.dataset.damageParts || "%5B%5D")); } catch {}

      if (damageParts.length > 0) {
        for (const part of damageParts) {
          const result = await handleDamage({
            dmgF: part.formula,
            isCrit: false,
            keyword: part.type ? `Damage (${part.type})` : "Damage",
            damageType: part.type,
            sheet: fakeSheet,
            dmgAdvantage: false,
            dmgDisadvantage: false,
          });
          if (result === false) break;
        }
      } else if (dmgF) {
        await handleDamage({
          dmgF,
          isCrit: false,
          keyword: damageType ? `Damage (${damageType})` : "Damage",
          damageType,
          sheet: fakeSheet,
          dmgAdvantage: false,
          dmgDisadvantage: false,
        });
      } else {
        ui.notifications.warn("No damage formula on this attack card.");
      }
    }
  });
});


