// modules/sheexcel_updated/sheexcel.js

// 1) Preload our Handlebars partials before anything renders
Hooks.once("init", async () => {
  const paths = [
    "modules/sheexcel_updated/templates/partials/main-tab.hbs",
    "modules/sheexcel_updated/templates/partials/references-tab.hbs",
    "modules/sheexcel_updated/templates/partials/settings-tab.hbs"
  ];
  console.log("⏳ Sheexcel loading templates from:", paths);
  try {
    await loadTemplates(paths);
    console.log("✅ Partials loaded:", Object.keys(Handlebars.partials));
  } catch (err) {
    console.error("❌ Sheexcel | loadTemplates failed:", err);
  }
});

import { prepareSheetData }   from "../helpers/prepareData.js";
import { importJsonHandler }  from "../helpers/importer.js";
import { batchFetchValues }   from "../helpers/batchFetcher.js";
import { handleRoll }         from "../helpers/roller.js";

export class SheexcelActorSheet extends ActorSheet {
  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      classes:   ["sheet", "actor", "sheexcel"],
      template:  "modules/sheexcel_updated/templates/sheet-template.html",
      width:     1200,
      height:    700,
      resizable: true,
      tabs: [
        { navSelector: ".sheexcel-sheet-tabs",      contentSelector: ".sheexcel-sidebar",       initial: "main"   },
        { navSelector: ".sheexcel-main-subtab-nav", contentSelector: ".sheexcel-main-subtabs", initial: "checks" }
      ]
    });
  }

  /** Override getData to merge in our flags */
  async getData(opts) {
    // 1) let Foundry build its normal data object…
    const data = await super.getData(opts);
    // 2) then augment with our flags, passing (data, actor)
    return prepareSheetData(data, this.actor);
  }

  /** Wire up all UI interactions */
  activateListeners(html) {
    super.activateListeners(html);

    // Toggle sidebar & primary tabs
    html.find(".sheexcel-sheet-toggle")
      .on("click", this._onToggleSidebar.bind(this));
    html.find(".sheexcel-sheet-main, .sheexcel-sheet-references, .sheexcel-sheet-settings")
      .on("click", this._onToggleTab.bind(this));

    // Update Sheet URL
    html.find(".sheexcel-setting-update-sheet")
      .on("click", this._onUpdateSheet.bind(this));

    // Reference rows: add / remove / fetch / save
    html.on("click", ".sheexcel-reference-add-button",    this._onAddReference.bind(this));
    html.on("click", ".sheexcel-reference-remove-button", this._onRemoveReference.bind(this));
    html.on("click", ".sheexcel-reference-remove-save",   this._onFetchAndUpdateCellValueByIndex.bind(this));
    html.on("click", ".sheexcel-reference-save-button",   this._onSaveReferences.bind(this));

    // Import JSON
    html.find(".sheexcel-import-json")
      .on("click", () => html.find("#sheexcel-json-file").click());
    html.find("#sheexcel-json-file")
      .on("change", e => importJsonHandler(e, this).then(() => this.render(false)));

    // Main subtabs and roll clicks
    html.find(".sheexcel-main-subtab-nav a.item")
      .on("click", this._onToggleSubtab.bind(this));
    html.find(".sheexcel-main-subtab-content")
      .on("click", ".sheexcel-roll", e => handleRoll(e, this));

    // Activate initial subtab
    html.find(".sheexcel-main-subtab-nav a.item.active").click();
  }

  /** Switch which main-subtab pane is visible */
  _onToggleSubtab(event) {
    event.preventDefault();
    const tab   = event.currentTarget.dataset.tab;
    this.element.find(".sheexcel-main-subtab-nav a.item").removeClass("active");
    this.element.find(".sheexcel-main-subtab-content").hide();
    $(event.currentTarget).addClass("active");
    this.element.find(`.sheexcel-main-subtab-content[data-tab="${tab}"]`).show();
  }

  /** Add a blank reference row in the DOM */
  _onAddReference(event) {
    event.preventDefault();
    const container = this.element.find(".sheexcel-references");
    const idx       = container.find(".sheexcel-reference-row").length;
    const sheets    = this.actor.getFlag("sheexcel_updated","sheetNames")||[];
    const options   = sheets.map(n => `<option value="${n}">${n}</option>`).join("");
    const row = $(`
      <div class="sheexcel-reference-row" data-index="${idx}">
        <input class="sheexcel-reference-input" data-type="cell"    data-index="${idx}" placeholder="Cell">
        <input class="sheexcel-reference-input" data-type="keyword" data-index="${idx}" placeholder="Keyword">
        <select class="sheexcel-reference-input" data-type="sheet"  data-index="${idx}">${options}</select>
        <select class="sheexcel-reference-input" data-type="refType" data-index="${idx}">
          <option value="checks">Checks</option>
          <option value="saves">Saves</option>
          <option value="attacks">Attacks</option>
          <option value="spells">Spells</option>
        </select>
        <div class="sheexcel-reference-remove">
          <button class="sheexcel-reference-remove-save"   data-index="${idx}">🔄</button>
          <button class="sheexcel-reference-remove-button" data-index="${idx}">Remove</button>
          <span class="sheexcel-reference-value"></span>
        </div>
      </div>`);
    container.append(row);
  }

  /** Remove one reference row from the DOM */
  _onRemoveReference(event) {
    event.preventDefault();
    $(event.currentTarget).closest(".sheexcel-reference-row").remove();
  }

  /** Fetch & update one cell’s value, then re-render */
  async _onFetchAndUpdateCellValueByIndex(index) {
    const refs    = foundry.utils.deepClone(await this.actor.getFlag("sheexcel_updated","cellReferences") || []);
    const sheetId = this.actor.getFlag("sheexcel_updated","sheetId");
    if (!refs[index] || !sheetId) return;

    const { cell, sheet } = refs[index];
    const safeSheet = sheet.match(/[^A-Za-z0-9_]/)
      ? `'${sheet.replace(/'/g,"''")}'`
      : sheet;
    const range = `${safeSheet}!${cell}`;
    try {
      const res  = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}?key=AIzaSyCAYacdw4aB7GtoxwnlpaF3aFZ2DgcJNHo`);
      if (!res.ok) throw new Error(res.statusText);
      const json = await res.json();
      refs[index].value = json.values?.[0]?.[0] || "";
    } catch {
      refs[index].value = "";
    }
    await this.actor.setFlag("sheexcel_updated","cellReferences",refs);
    this.render(false);
  }

  /** Read rows, batch-fetch every cell, then persist & re-render */
  async _onSaveReferences(event) {
    event.preventDefault();
    // Collect all inputs including extra attack fields
    const rows = this.element.find(".sheexcel-reference-row").toArray();
    const refs = rows.map(el => {
      const $r = $(el);
      return {
        cell:           $r.find("input[data-type='cell']").val().trim(),
        keyword:        $r.find("input[data-type='keyword']").val().trim(),
        sheet:          $r.find("select[data-type='sheet']").val(),
        type:           $r.find("select[data-type='refType']").val(),
        attackNameCell: $r.find("input[data-type='attackNameCell']").val()?.trim() || "",
        critRangeCell:  $r.find("input[data-type='critRangeCell']").val()?.trim()  || "",
        damageCell:     $r.find("input[data-type='damageCell']").val()?.trim()     || "",
        value:          ""
      };
    });

    const sheetId = this.actor.getFlag("sheexcel_updated","sheetId");
    if (sheetId) {
      const updated = await batchFetchValues(sheetId, refs);
      await this.actor.setFlag("sheexcel_updated","cellReferences",updated);
    }
    this.render(false);
  }

  /** Toggle the sidebar collapsed state */
  _onToggleSidebar(event) {
    event.preventDefault();
    const c = !this.actor.getFlag("sheexcel_updated","sidebarCollapsed");
    this.actor.setFlag("sheexcel_updated","sidebarCollapsed",c).then(() => this.render(false));
  }

  /** Switch primary sidebar tabs */
  _onToggleTab(event) {
    event.preventDefault();
    const tab = event.currentTarget.dataset.tab;
    this.element.find(".sheexcel-sheet-tabs .item, .sheexcel-sidebar-tab").removeClass("active");
    this.element.find(`.sheexcel-sidebar-tab[data-tab="${tab}"]`).addClass("active");
    event.currentTarget.classList.add("active");
  }

  /** Fetch sheet metadata & store URL, ID, sheetNames */
  _onUpdateSheet(event) {
    event.preventDefault();
    const url = this.element.find("#sheexcel-setting-url").val()?.trim();
    if (!url) return ui.notifications.warn("Enter a valid Google Sheet URL.");
    const match   = url.match(/\/d\/([^\/]+)/);
    const sheetId = match?.[1];
    if (!sheetId) return ui.notifications.error("Couldn’t extract Sheet ID.");
    fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties.title&key=AIzaSyCAYacdw4aB7GtoxwnlpaF3aFZ2DgcJNHo`)
      .then(r => r.ok ? r.json() : Promise.reject(r.statusText))
      .then(json => {
        const names = json.sheets.map(s => s.properties.title);
        return Promise.all([
          this.actor.setFlag("sheexcel_updated","sheetUrl",url),
          this.actor.setFlag("sheexcel_updated","sheetId",sheetId),
          this.actor.setFlag("sheexcel_updated","sheetNames",names)
        ]);
      })
      .then(() => this.render(false))
      .catch(() => ui.notifications.error("Failed to load sheet metadata."));
  }
}

// Register our sheet
Hooks.once("setup", () => {
  Actors.registerSheet("sheexcel", SheexcelActorSheet, {
    types: ["character","npc","creature","vehicle"],
    label: "Sheexcel",
    makeDefault: false
  });
});

// Wrap prepareDerivedData so our sheexcel values appear on actor.system.sheexcel
Hooks.once("ready", () => {
  libWrapper.register(
    "sheexcel_updated",
    "CONFIG.Actor.documentClass.prototype.prepareDerivedData",
    (wrapped, ...args) => {
      wrapped(...args);
      const refs = this.getFlag("sheexcel_updated","cellReferences") || [];
      this.system = this.system || {};
      this.system.sheexcel = refs.reduce((o,r) => {
        if (r.keyword) o[r.keyword] = r.value;
        return o;
      }, {});
    },
    "WRAPPER"
  );
});
