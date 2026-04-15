import { MODULE_NAME, SETTINGS } from "../helpers/constants.js";
import { handleInlineRollInteraction } from "../helpers/inlineRolls.js";

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

  // Roll mode setting (client)
  game.settings.register("sheexcel", SETTINGS.ROLL_MODE, {
    name: "Default Roll Mode", 
    scope: "client",
    config: false,
    type: String,
    default: "norm"
  });

  // Damage modes setting (client)
  game.settings.register("sheexcel", SETTINGS.DAMAGE_MODES, {
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
  Actors.registerSheet("sheexcel", SheexcelActorSheet, {
    types: ["character", "npc", "creature", "vehicle"],
    label: "Sheexcel",
    makeDefault: false
  });
});

// --- Derived Data Wrapper ---
Hooks.once("ready", () => {
  libWrapper.register(
    "sheexcel_updated",
    "CONFIG.Actor.documentClass.prototype.prepareDerivedData",
    function(wrapped, ...args) {
      wrapped(...args);
      const refs = this.getFlag("sheexcel_updated", "cellReferences") || [];
      this.system = this.system || {};
      this.system.sheexcel = refs.reduce((o, r) => {
        if (r.keyword) o[r.keyword] = r.value;
        return o;
      }, {});
    },
    "WRAPPER"
  );

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

// --- Sheet Resizer ---
Hooks.on("renderActorSheet", (app, html, data) => {
  const wrapper = html.find('.sheexcel-sheet-google-wrapper');
  const resizer = html.find('.sheexcel-sheet-resizer');
  let isResizing = false;
  let startY, startHeight;

  resizer.on('mousedown', function (e) {
    isResizing = true;
    startY = e.clientY;
    startHeight = wrapper.height();
    $(document).on('mousemove.sheexcelResize', function (e) {
      if (!isResizing) return;
      let newHeight = Math.max(100, startHeight + (e.clientY - startY));
      wrapper.height(newHeight);
    });
    $(document).on('mouseup.sheexcelResize', function () {
      isResizing = false;
      $(document).off('.sheexcelResize');
    });
    e.preventDefault();
  });
});