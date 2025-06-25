Hooks.once("init", async function () {

  for (const type of ["character", "npc", "creature", "vehicle"]) {
    if (!CONFIG.Actor.sheetClasses[type]) CONFIG.Actor.sheetClasses[type] = {};
    CONFIG.Actor.sheetClasses[type]["sheexcel"] = {
      label: "Sheexcel",
      sheetClass: SheexcelActorSheet,
      makeDefault: false
    };
  }

  libWrapper.register("sheexcel", "CONFIG.Actor.documentClass.prototype.prepareData", function (wrapped) {
    wrapped.call(this);

    const refs = this.getFlag("sheexcel", "cellReferences") || [];
    const sheexcel = {};

    for (const ref of refs) {
      if (ref.keyword && ref.value !== undefined) {
        sheexcel[ref.keyword] = ref.value;
      }
    }

    if (!this.system) this.system = {};
    this.system.sheexcel = sheexcel;
  }, "WRAPPER");

  game.sheexcel = {
    getSheexcelValue: (actorId, keyword) => {
      const actor = game.actors.get(actorId);
      if (actor && actor.sheet?.getSheexcelValue) {
        return {
          value: actor.sheet.getSheexcelValue(keyword),
          actor: actor
        };
      }
      return null;
    }
  };

});

class SheexcelActorSheet extends ActorSheet {
  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      classes: ["sheet", "actor", "sheexcel"],
      template: "modules/sheexcel/templates/sheet-template.html",
      width: 1200,
      height: 700,
      resizable: true,
      tabs: [{
        navSelector: ".sheexcel-sheet-tabs",
        contentSelector: ".sheexcel-sidebar",
        initial: "settings"
      }]
    });
  }

  async getData() {
    const data = await super.getData();
    const flags = this.actor.flags.sheexcel || {};
    data.sheetUrl = flags.sheetUrl || "";
    data.zoomLevel = flags.zoomLevel || 100;
    data.hideMenu = flags.hideMenu ?? true;
    data.sidebarCollapsed = flags.sidebarCollapsed || false;
    data.cellReferences = foundry.utils.deepClone(flags.cellReferences || []);
    data.sheetNames = flags.sheetNames?.length > 1 ? flags.sheetNames : null;
    data.currentSheetName = flags.currentSheetName || null;
    data.sheetId = flags.sheetId || null;

    data.adjustedReferences = (data.cellReferences || []).map(ref => {
      const sheetMap = {};
      (data.sheetNames || []).forEach(name => sheetMap[name] = name);
      ref.sheetNames = sheetMap;
      if (typeof ref.value === 'string' && ref.value.length > 10) {
        ref.value = ref.value.slice(0, 10);
      }
      return ref;
    });

    return data;
  }

  activateListeners(html) {
    super.activateListeners(html);
    html.find(".sheexcel-sheet-toggle").on("click", this._onToggleSidebar.bind(this));
    html.find(".sheexcel-sheet-references").on("click", this._onToggleTab.bind(this));
    html.find(".sheexcel-sheet-settings").on("click", this._onToggleTab.bind(this));
    html.find(".sheexcel-setting-update-sheet").on("click", this._onUpdateSheet.bind(this));
    html.find(".sheexcel-reference-add-button").on("click", this._onAddReference.bind(this));
    html.on("click", ".sheexcel-reference-remove-save", this._onSaveReference.bind(this));
    html.on("click", ".sheexcel-reference-remove-button", this._onRemoveReference.bind(this));
    html.on("change", "#sheexcel-cell", this._onCellReferenceChange.bind(this));
    html.on("change", "#sheexcel-keyword", this._onKeywordReferenceChange.bind(this));
    html.on("change", "#sheexcel-sheet", this._onCellReferenceChange.bind(this));

    this._iframe = html.find(".sheexcel-iframe")[0];
    this._setupZoom(html);
    this._setupHideMenu(html);
    this._applyZoom();
  }

  async _onUpdateSheet(event) {
    event.preventDefault();
    const sheetUrl = this.element.find('input[name="sheetUrl"]').val();
    if (!sheetUrl) {
      await this.actor.update({
        'flags.sheexcel.sheetId': null,
        'flags.sheexcel.currentSheetName': null,
        'flags.sheexcel.sheetNames': [],
        'flags.sheexcel.sheetUrl': ""
      });
      return this.render();
    }
    const sheetIdMatch = sheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!sheetIdMatch) return ui.notifications.error("Invalid Google Sheet URL");
    const sheetId = sheetIdMatch[1];
    const sheetNames = await this._fetchSheetNames(sheetId);
    const currentSheetName = sheetNames?.[0] || null;
    await this.actor.update({
      'flags.sheexcel.sheetId': sheetId,
      'flags.sheexcel.sheetNames': sheetNames,
      'flags.sheexcel.currentSheetName': currentSheetName,
      'flags.sheexcel.sheetUrl': sheetUrl
    });
    this.render();
  }

  async _fetchSheetNames(sheetId) {
    const url = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
    try {
      const response = await fetch(url);
      const html = await response.text();
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const tabs = doc.querySelectorAll('.docs-sheet-tab');
      return Array.from(tabs).map(tab => tab.textContent.trim()).filter(Boolean);
    } catch (e) {
      console.error("Failed to fetch sheet names:", e);
      return [];
    }
  }

  async _onAddReference(event) {
    event.preventDefault();
    const flags = this.actor.flags.sheexcel || {};
    const refs = foundry.utils.deepClone(flags.cellReferences || []);
    refs.push({ sheet: flags.currentSheetName, cell: "", keyword: "", value: "" });
    await this.actor.setFlag("sheexcel", "cellReferences", refs);
    this.render();
  }

  async _onRemoveReference(event) {
    event.preventDefault();
    const index = $(event.currentTarget).closest(".sheexcel-reference-cell").index();
    const refs = foundry.utils.deepClone(this.actor.getFlag("sheexcel", "cellReferences"));
    refs.splice(index, 1);
    await this.actor.setFlag("sheexcel", "cellReferences", refs);
    this.render();
  }

  async _onSaveReference(event) {
    event.preventDefault();
    await this._refetchAllCellValues();
  }

  async _refetchAllCellValues() {
    const flags = this.actor.flags.sheexcel || {};
    const refs = foundry.utils.deepClone(flags.cellReferences || []);
    const updated = [];
    for (const ref of refs) {
      if (ref.cell && ref.sheet) {
        ref.value = await this._fetchCellValue(flags.sheetId, ref.sheet, ref.cell);
      }
      updated.push(ref);
    }
    await this.actor.setFlag("sheexcel", "cellReferences", updated);
    this.render();
  }

  async _fetchCellValue(sheetId, sheetName, cell) {
    if (!sheetId || !sheetName || !cell) return "";
    const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}&range=${cell}`;
    try {
      const res = await fetch(url);
      const text = (await res.text()).trim();
      return text.replace(/^"|"$/g, "");
    } catch (e) {
      console.warn("Failed to fetch cell value:", e);
      return "";
    }
  }

  async _onCellReferenceChange(event) {
    const $el = $(event.currentTarget);
    const index = $el.closest(".sheexcel-reference-cell").index();
    const refs = foundry.utils.deepClone(await this.actor.getFlag("sheexcel", "cellReferences"));
    const field = $el.attr("id").replace("sheexcel-", "");
    refs[index][field] = $el.val();
    await this.actor.setFlag("sheexcel", "cellReferences", refs);
    await this._refetchAllCellValues();
  }

  async _onKeywordReferenceChange(event) {
    const index = $(event.currentTarget).closest(".sheexcel-reference-cell").index();
    const refs = foundry.utils.deepClone(await this.actor.getFlag("sheexcel", "cellReferences"));
    refs[index].keyword = event.currentTarget.value;
    await this.actor.setFlag("sheexcel", "cellReferences", refs);
  }

  _setupZoom(html) {
    const slider = html.find("#sheexcel-setting-zoom-slider")[0];
    const display = html.find("#sheexcel-setting-zoom-value")[0];
    if (!slider || !display) return;
    slider.addEventListener("input", async e => {
      const zoom = parseInt(e.target.value);
      display.textContent = `${zoom}%`;
      await this.actor.setFlag("sheexcel", "zoomLevel", zoom);
      this._applyZoom();
    });
  }

  _applyZoom() {
    const zoom = this.actor.getFlag("sheexcel", "zoomLevel") || 100;
    if (this._iframe) {
      this._iframe.style.transform = `scale(${zoom / 100})`;
      this._iframe.style.transformOrigin = "top left";
      this._iframe.style.width = `${100 * (100 / zoom)}%`;
      this._iframe.style.height = `${100 * (100 / zoom)}%`;
    }
  }

  _setupHideMenu(html) {
    const checkbox = html.find("#sheexcel-setting-hide-menu")[0];
    if (!checkbox) return;
    checkbox.addEventListener("change", async e => {
      const hide = e.target.checked;
      await this.actor.setFlag("sheexcel", "hideMenu", hide);
      this._updateIframeSrc(hide);
    });
  }

  _updateIframeSrc(hideMenu) {
    if (!this._iframe) return;
    const url = this.actor.getFlag("sheexcel", "sheetUrl");
    if (!url) return;
    const rm = hideMenu ? "minimal" : "embedded";
    this._iframe.src = `${url}?embedded=true&rm=${rm}`;
  }

  getSheexcelValue(keyword) {
    return this.actor.system?.sheexcel?.[keyword] || null;
  }
}