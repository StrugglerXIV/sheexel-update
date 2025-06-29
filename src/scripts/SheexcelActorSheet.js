import { prepareSheetData } from "../helpers/prepareData.js";
import { randomID } from "../helpers/idGenerator.js";
import { attachSheexcelListeners } from "../helpers/sheexcelListeners.js";


// --- Main Sheet Class ---
export class SheexcelActorSheet extends ActorSheet {
  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      classes: ["sheet", "actor", "sheexcel"],
      template: "modules/sheexcel_updated/templates/sheet-template.html",
      width: 1200,
      height: 700,
      resizable: true,
      tabs: [
        { navSelector: ".sheexcel-sheet-tabs", contentSelector: ".sheexcel-sidebar", initial: "main" },
        { navSelector: ".sheexcel-main-subtab-nav", contentSelector: ".sheexcel-main-subtabs", initial: "checks" }
      ]
    });
  }

  async getData(opts) {
    const data = await super.getData(opts);
    return prepareSheetData(data, this.actor);
  }

  _onRefTypeChange(event) {
    const $select = $(event.currentTarget);
    const $row = $select.closest(".sheexcel-reference-row");
    $row.find(".sheexcel-attack-fields").remove();
    if ($select.val() === "attacks") {
      const idx = $row.data("index");
      const attackFields = $(`
        <div class="sheexcel-attack-fields">
          <input class="sheexcel-reference-input" data-type="attackNameCell" data-index="${idx}" placeholder="Attack Name Cell">
          <input class="sheexcel-reference-input" data-type="critRangeCell" data-index="${idx}" placeholder="Crit Range Cell">
          <input class="sheexcel-reference-input" data-type="damageCell" data-index="${idx}" placeholder="Damage Cell">
        </div>
      `);
      $row.find("select[data-type='refType']").after(attackFields);
    }
  }

  // Activate listeners 
    activateListeners(html) {
    super.activateListeners(html);
    attachSheexcelListeners(this, html);
  }



  // --- Reference Row Management ---
  _onAddReference(event) {
    event.preventDefault();
    const container = this.element.find(".sheexcel-references");
    const idx = container.find(".sheexcel-reference-row").length;
    const sheets = this.actor.getFlag("sheexcel_updated", "sheetNames") || [];
    const options = sheets.map(n => `<option value="${n}">${n}</option>`).join("");
    const newId = randomID();
    const row = $(`
      <div class="sheexcel-reference-row" data-index="${idx}" data-id="${newId}">
        <input class="sheexcel-reference-input" data-type="cell"    data-index="${idx}" placeholder="Bonus">
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

  _onRemoveReference(event) {
    event.preventDefault();
    $(event.currentTarget).closest(".sheexcel-reference-row").remove();
  }

  async _onFetchAndUpdateCellValueByIndex(index) {
    const refs = foundry.utils.deepClone(this.actor.getFlag("sheexcel_updated", "cellReferences") || []);
    const sheetId = this.actor.getFlag("sheexcel_updated", "sheetId");
    if (!refs[index] || !sheetId) return;
    const { cell, sheet } = refs[index];
    const safeSheet = sheet.match(/[^A-Za-z0-9_]/)
      ? `'${sheet.replace(/'/g, "''")}'`
      : sheet;
    const range = `${safeSheet}!${cell}`;
    try {
      const apiKey = game.settings.get("sheexcel_updated", "googleApiKey");
      const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}?key=${apiKey}`);
      if (!res.ok) throw new Error(res.statusText);
      const json = await res.json();
      refs[index].value = json.values?.[0]?.[0] || "";
    } catch {
      refs[index].value = "";
    }
    await this.actor.setFlag("sheexcel_updated", "cellReferences", refs);
    this.render(false);
  }


  _onToggleSidebar(event) {
    event.preventDefault();
    const c = !this.actor.getFlag("sheexcel_updated", "sidebarCollapsed");
    this.element.find('.sheexcel-sidebar').toggleClass('collapsed', c);
    this.actor.setFlag("sheexcel_updated", "sidebarCollapsed", c);
  }

  _onToggleTab(event) {
    event.preventDefault();
    const tab = event.currentTarget.dataset.tab;
    this.element.find(".sheexcel-sheet-tabs .item, .sheexcel-sidebar-tab").removeClass("active");
    this.element.find(`.sheexcel-sidebar-tab[data-tab="${tab}"]`).addClass("active");
    event.currentTarget.classList.add("active");
  }

  _onUpdateSheet(event) {
    event.preventDefault();
    const url = this.element.find("#sheexcel-setting-url").val()?.trim();
    if (!url) return ui.notifications.warn("Enter a valid Google Sheet URL.");
    const match = url.match(/\/d\/([^\/]+)/);
    const sheetId = match?.[1];
    if (!sheetId) return ui.notifications.error("Couldn’t extract Sheet ID.");
    const apiKey = game.settings.get("sheexcel_updated", "googleApiKey");
    fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties.title&key=${apiKey}`)
      .then(r => r.ok ? r.json() : Promise.reject(r.statusText))
      .then(json => {
        const names = json.sheets.map(s => s.properties.title);
        return Promise.all([
          this.actor.setFlag("sheexcel_updated", "sheetUrl", url),
          this.actor.setFlag("sheexcel_updated", "sheetId", sheetId),
          this.actor.setFlag("sheexcel_updated", "sheetNames", names)
        ]);
      })
      .then(() => this.render(false))
      .catch(() => ui.notifications.error("Failed to load sheet metadata."));
  }

  _onToggleSubtab(event) {
    event.preventDefault();
    const tab = event.currentTarget.dataset.tab;
    this.element.find(".sheexcel-main-subtab-nav a.item").removeClass("active");
    this.element.find(".sheexcel-main-subtab-content").hide();
    $(event.currentTarget).addClass("active");
    this.element.find(`.sheexcel-main-subtab-content[data-tab="${tab}"]`).show();
  }
}