import { MODULE_NAME, SETTINGS } from "../helpers/constants.js";

// 1) Preload our Handlebars partials before anything renders
Hooks.once("init", async () => {
  console.log("🔄 Sheexcel | Module initialization starting...");
  
  const paths = [
    "modules/sheexcel_updated/templates/partials/main-tab.hbs",
    "modules/sheexcel_updated/templates/partials/references-tab.hbs",
    "modules/sheexcel_updated/templates/partials/settings-tab.hbs"
  ];
  try {
    await loadTemplates(paths);
    console.log("✅ Sheexcel | Templates loaded successfully");
  } catch (err) {
    console.error("❌ Sheexcel | loadTemplates failed:", err);
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

  // Google OAuth Client ID (for write access)
  game.settings.register(MODULE_NAME, SETTINGS.GOOGLE_OAUTH_CLIENT_ID, {
    name: "Google OAuth Client ID",
    hint: "OAuth Client ID for write access (Google Identity Services).",
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