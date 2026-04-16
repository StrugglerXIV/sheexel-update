import { prepareSheetData } from "../helpers/prepareData.js";
import { randomID } from "../helpers/idGenerator.js";
import { attachSheexcelListeners } from "../helpers/sheexcelListeners.js";
import { MODULE_NAME, FLAGS, SETTINGS, CSS_CLASSES, ERROR_MESSAGES, API_CONFIG } from "../helpers/constants.js";
import { validateSheetUrl, validateCellReference, ValidationError } from "../helpers/validation.js";
import { loadingManager } from "../helpers/loadingManager.js";
import { apiCache } from "../helpers/apiCache.js";
import { enrichTextWithInlineRolls, escapeHtml } from "../helpers/inlineRolls.js";


// --- Main Sheet Class ---
export class SheexcelActorSheet extends ActorSheet {
  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      classes: ["sheet", "actor", CSS_CLASSES.SHEET],
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

  async _render(force = false, options = {}) {
    // Preserve key scroll containers across rerenders to prevent jump-to-top on input updates.
    const scrollSnapshot = {};
    const hasElement = Boolean(this.element?.length);

    if (hasElement) {
      const selectors = [
        ".window-content",
        "form.sheexcel-sheet",
        ".sheexcel-sidebar",
        ".sheexcel-sidebar-tab.active",
        ".sheexcel-main-subtabs",
        ".sheexcel-main-subtab-content:visible",
        ".sheexcel-references-list"
      ];

      for (const selector of selectors) {
        const el = this.element.find(selector)[0];
        if (!el) continue;
        if (el.scrollHeight <= el.clientHeight) continue;
        scrollSnapshot[selector] = el.scrollTop;
      }
    }

    await super._render(force, options);

    const restoreScroll = () => {
      for (const [selector, top] of Object.entries(scrollSnapshot)) {
        const el = this.element?.find(selector)[0];
        if (el) el.scrollTop = top;
      }
    };

    // Restore immediately, then once more next frame in case Foundry applies late layout updates.
    restoreScroll();
    requestAnimationFrame(restoreScroll);
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

    try {
      attachSheexcelListeners(this, html);
    } catch (error) {
      console.error("Sheexcel | Listener initialization failed:", error);
    }

    try {
      this._startOrbPolling(html);
    } catch (error) {
      console.error("Sheexcel | Orb initialization failed:", error);
    }

    this._loadArmorFromFlags(html);
    this._loadStatsFromFlags(html);
  }

  async close(options) {
    if (this._orbPoller) {
      clearInterval(this._orbPoller);
      this._orbPoller = null;
    }
    return super.close(options);
  }



  // --- Reference Row Management ---
  _columnLettersToNumber(letters) {
    let value = 0;
    const cleaned = String(letters || "").trim().toUpperCase();
    for (let i = 0; i < cleaned.length; i++) {
      const code = cleaned.charCodeAt(i);
      if (code < 65 || code > 90) throw new Error("Invalid start column");
      value = value * 26 + (code - 64);
    }
    if (!value) throw new Error("Invalid start column");
    return value;
  }

  _columnNumberToLetters(number) {
    let n = Number(number);
    if (!Number.isInteger(n) || n < 1) throw new Error("Invalid column number");
    let letters = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      letters = String.fromCharCode(65 + rem) + letters;
      n = Math.floor((n - 1) / 26);
    }
    return letters;
  }

  _buildRange(startColumn, startOffset, endOffset, row) {
    const start = this._columnNumberToLetters(startColumn + startOffset);
    const end = this._columnNumberToLetters(startColumn + endOffset);
    return `${start}${row}:${end}${row}`;
  }

  _buildAbsoluteRange(startColumn, endColumn, row) {
    const start = this._columnNumberToLetters(startColumn);
    const end = this._columnNumberToLetters(endColumn);
    return `${start}${row}:${end}${row}`;
  }

  _parseA1Cell(value) {
    const raw = String(value || "").trim().toUpperCase();
    const match = raw.match(/^([A-Z]+)(\d+)$/);
    if (!match) throw new Error(`Invalid cell: ${value}`);
    return {
      column: this._columnLettersToNumber(match[1]),
      row: Number(match[2])
    };
  }

  _buildSingleCell(column, row) {
    return `${this._columnNumberToLetters(column)}${row}`;
  }

  _buildColumnRange(column, startRow, endRow) {
    const col = this._columnNumberToLetters(column);
    return `${col}${startRow}:${col}${endRow}`;
  }

  _buildRectRange(startColumn, startRow, endColumn, endRow) {
    const startCol = this._columnNumberToLetters(startColumn);
    const endCol = this._columnNumberToLetters(endColumn);
    return `${startCol}${startRow}:${endCol}${endRow}`;
  }

  _invalidateApiCache() {
    const sheetId = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_ID);
    if (sheetId) {
      apiCache.invalidateSheet(sheetId);
    }
  }

  _sanitizeSheetName(sheet) {
    return sheet.match(/[^A-Za-z0-9_]/)
      ? `'${sheet.replace(/'/g, "''")}'`
      : sheet;
  }

  async _scanArea(sheet, topLeftCell, bottomRightCell) {
    const topLeft = this._parseA1Cell(topLeftCell);
    const bottomRight = this._parseA1Cell(bottomRightCell);

    const startColumn = Math.min(topLeft.column, bottomRight.column);
    const endColumn = Math.max(topLeft.column, bottomRight.column);
    const startRow = Math.min(topLeft.row, bottomRight.row);
    const endRow = Math.max(topLeft.row, bottomRight.row);

    if (startRow === endRow || startColumn === endColumn) {
      throw new Error("Range must contain multiple rows and columns");
    }

    const safeSheet = this._sanitizeSheetName(sheet);

    const areaRange = `${safeSheet}!${this._columnNumberToLetters(startColumn)}${startRow}:${this._columnNumberToLetters(endColumn)}${endRow}`;
    const sheetId = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_ID);
    if (!sheetId) throw new Error("Missing sheet ID. Update Sheet first.");

    const scanJson = await apiCache.batchGet(sheetId, [areaRange]);
    const matrix = scanJson.valueRanges?.[0]?.values || [];

    const readCell = (absoluteRow, absoluteColumn) => {
      const rowOffset = absoluteRow - startRow;
      const colOffset = absoluteColumn - startColumn;
      return (matrix[rowOffset]?.[colOffset] ?? "").toString().trim();
    };

    return { startColumn, endColumn, startRow, endRow, readCell };
  }

  async _fetchCellNotes(sheet, topLeftCell, bottomRightCell) {
    const topLeft = this._parseA1Cell(topLeftCell);
    const bottomRight = this._parseA1Cell(bottomRightCell);

    const startColumn = Math.min(topLeft.column, bottomRight.column);
    const endColumn = Math.max(topLeft.column, bottomRight.column);
    const startRow = Math.min(topLeft.row, bottomRight.row);
    const endRow = Math.max(topLeft.row, bottomRight.row);

    const safeSheet = this._sanitizeSheetName(sheet);
    const areaRange = `${safeSheet}!${this._columnNumberToLetters(startColumn)}${startRow}:${this._columnNumberToLetters(endColumn)}${endRow}`;

    const sheetId = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_ID);
    const apiKey = game.settings.get(MODULE_NAME, SETTINGS.GOOGLE_API_KEY);
    if (!sheetId) throw new Error("Missing sheet ID. Update Sheet first.");
    if (!apiKey) throw new Error("Missing Google API key.");

    const url = `${API_CONFIG.BASE_URL}/${sheetId}?ranges=${encodeURIComponent(areaRange)}&includeGridData=true&fields=sheets(data(rowData(values(note))))&key=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Notes fetch failed (${response.status})`);
    const json = await response.json();
    const rows = json.sheets?.[0]?.data?.[0]?.rowData || [];

    const notes = new Map();
    rows.forEach((rowData, rIdx) => {
      const values = rowData?.values || [];
      values.forEach((cell, cIdx) => {
        const note = cell?.note;
        if (!note) return;
        const absoluteRow = startRow + rIdx;
        const absoluteCol = startColumn + cIdx;
        notes.set(`${absoluteRow},${absoluteCol}`, note);
      });
    });

    return notes;
  }

  async _locateStatCells(sheetName, labels) {
    const scan = await this._scanArea(sheetName, "A1", "Z40");
    const normalize = (value) => String(value || "").trim().replace(/:$/, "").toLowerCase();

    let statRef = null;
    for (let row = scan.startRow; row <= scan.endRow; row++) {
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const text = normalize(scan.readCell(row, col));
        if (labels.includes(text)) {
          statRef = { row, col };
          break;
        }
      }
      if (statRef) break;
    }

    if (!statRef) return null;

    let headerRow = null;
    for (let row = statRef.row - 1; row >= Math.max(scan.startRow, statRef.row - 12); row--) {
      let found = false;
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const text = normalize(scan.readCell(row, col));
        if (text === "base" || text === "tot" || text === "total") {
          found = true;
          break;
        }
      }
      if (found) {
        headerRow = row;
        break;
      }
    }

    if (!headerRow) return null;

    let baseCol = null;
    let totalCol = null;
    for (let col = scan.startColumn; col <= scan.endColumn; col++) {
      const text = normalize(scan.readCell(headerRow, col));
      if (text === "base") baseCol = col;
      if (text === "tot" || text === "total") totalCol = col;
    }

    if (!baseCol || !totalCol) return null;

    return {
      sheetName,
      baseCell: this._buildSingleCell(baseCol, statRef.row),
      totalCell: this._buildSingleCell(totalCol, statRef.row)
    };
  }

  async _getStatValues(labels, cacheKey, options = {}) {
    const { allowZeroMax = false } = options;
    const cells = await this._getStatCells(labels, cacheKey);
    if (!cells) return null;

    const safeSheet = this._sanitizeSheetName(cells.sheetName);
    const baseRange = `${safeSheet}!${cells.baseCell}`;
    const totalRange = `${safeSheet}!${cells.totalCell}`;
    const response = await this._fetchValuesNoCache(cells.sheetId, [baseRange, totalRange]);
    const baseRaw = response.valueRanges?.[0]?.values?.[0]?.[0] ?? "";
    const totalRaw = response.valueRanges?.[1]?.values?.[0]?.[0] ?? "";

    let max = this._parseRawNumber(baseRaw);
    let current = this._parseRawNumber(totalRaw);
    if (allowZeroMax) {
      if (!Number.isFinite(max)) max = 0;
      if (!Number.isFinite(current)) current = 0;
    }
    if (!Number.isFinite(max) || !Number.isFinite(current)) return null;

    return { current, max };
  }

  async _fetchValuesNoCache(sheetId, ranges) {
    const apiKey = game.settings.get(MODULE_NAME, SETTINGS.GOOGLE_API_KEY);
    if (!apiKey) throw new Error("Missing Google API key.");

    const params = new URLSearchParams();
    ranges.forEach(range => params.append("ranges", range));
    params.set("valueRenderOption", "UNFORMATTED_VALUE");
    params.set("key", apiKey);

    const url = `${API_CONFIG.BASE_URL}/${sheetId}/values:batchGet?${params.toString()}`;
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HP fetch failed (${response.status})`);
    return response.json();
  }

  async _locateArmorCells(sheetName) {
    const scan = await this._scanArea(sheetName, "A1", "Z80");
    const normalize = (value) => String(value || "").trim().replace(/:$/, "").toLowerCase();
    const labels = ["blunt", "slash", "pierce", "acid", "fire", "cold", "lightning", "necrotic", "arcane", "psychic", "radiant", "raidant"];

    let armorRef = null;
    for (let row = scan.startRow; row <= scan.endRow; row++) {
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        if (normalize(scan.readCell(row, col)) === "armor") {
          armorRef = { row, col };
          break;
        }
      }
      if (armorRef) break;
    }

    if (!armorRef) return null;

    const entriesByLabel = new Map();
    const maxRow = Math.min(scan.endRow, armorRef.row + 20);
    for (let row = armorRef.row + 1; row <= maxRow; row++) {
      const raw = String(scan.readCell(row, armorRef.col) || "").trim();
      if (!raw) continue;
      const normalized = normalize(raw);
      const matched = labels.find(label => normalized.includes(label));
      if (!matched) continue;

      const hasNumber = /[+-]?\d+(?:\.\d+)?/.test(raw);
      const cellRef = this._buildSingleCell(armorRef.col, row);
      const existing = entriesByLabel.get(matched);
      if (!existing || (hasNumber && !existing.hasNumber)) {
        entriesByLabel.set(matched, { label: matched, cell: cellRef, hasNumber });
      }
    }

    const entries = Array.from(entriesByLabel.values()).map(({ label, cell }) => ({ label, cell }));

    if (!entries.length) return null;

    return { sheetName, entries };
  }

  async _getArmorCells(options = {}) {
    const { force = false } = options;
    const sheetId = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_ID);
    if (!sheetId) return null;

    if (this._armorCells?.sheetName) {
      return { sheetId, ...this._armorCells };
    }

    const now = Date.now();
    if (!force && this._armorLocateAttempt && (now - this._armorLocateAttempt) < 60000) {
      return null;
    }
    this._armorLocateAttempt = now;

    const metadata = await this._fetchSheetMetadata();
    const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
    if (!sheets.length) return null;

    const coreSheet = sheets.find(s => /^core$/i.test(s.properties?.title || ""))
      || sheets.find(s => /core/i.test(s.properties?.title || ""));
    if (!coreSheet?.properties?.title) return null;

    const sheetName = coreSheet.properties.title;
    const located = await this._locateArmorCells(sheetName);
    if (!located) return null;

    this._armorCells = located;
    return { sheetId, ...located };
  }

  async _updateArmorBlock(html, options = {}) {
    const { force = false } = options;
    if (this._armorUpdateInFlight) return;
    this._armorUpdateInFlight = true;
    try {
      const cells = await this._getArmorCells({ force });
      if (!cells) return;

      const safeSheet = this._sanitizeSheetName(cells.sheetName);
      const ranges = cells.entries.map(entry => `${safeSheet}!${entry.cell}`);
      const response = await this._fetchValuesNoCache(cells.sheetId, ranges);
      const values = response.valueRanges || [];

      const normalizeRange = (range) => {
        const raw = String(range || "");
        const parts = raw.split("!");
        return parts.length > 1 ? parts[1].replace(/\$/g, "") : raw.replace(/\$/g, "");
      };

      const valueByCell = new Map();
      values.forEach((rangeEntry) => {
        const cellKey = normalizeRange(rangeEntry?.range);
        if (!cellKey) return;
        const raw = String(rangeEntry?.values?.[0]?.[0] ?? "").trim();
        valueByCell.set(cellKey, raw);
      });

      const valueMap = new Map();
      cells.entries.forEach((entry, index) => {
        const raw = valueByCell.get(entry.cell) ?? String(values[index]?.values?.[0]?.[0] ?? "").trim();
        const labelPattern = new RegExp(`^\\s*([+-]?\\d+(?:\\.\\d+)?)\\s+${entry.label}\\b`, "i");
        const match = raw.match(labelPattern);
        const numberMatch = raw.match(/[+-]?\d+(?:\.\d+)?/);
        const value = match?.[1] ?? numberMatch?.[0] ?? "-";
        valueMap.set(entry.label, value);
      });

      const rows = html.find('.sheexcel-armor-row');
      rows.each((_, row) => {
        const label = String(row.dataset.armorLabel || "").trim().toLowerCase();
        const alt = label === "radiant" ? "raidant" : label;
        const value = valueMap.get(label) ?? valueMap.get(alt) ?? "-";
        $(row).find('.sheexcel-armor-value').text(value);
      });

      await this.actor.setFlag(MODULE_NAME, FLAGS.ARMOR_CACHE, {
        values: Object.fromEntries(valueMap),
        updatedAt: Date.now()
      });
    } finally {
      this._armorUpdateInFlight = false;
    }
  }

  _startArmorPolling(html) {
    this._updateArmorBlock(html).catch(() => {});
  }

  _loadArmorFromFlags(html) {
    const cache = this.actor.getFlag(MODULE_NAME, FLAGS.ARMOR_CACHE);
    if (!cache?.values) return;
    const rows = html.find('.sheexcel-armor-row');
    rows.each((_, row) => {
      const label = String(row.dataset.armorLabel || "").trim().toLowerCase();
      const alt = label === "radiant" ? "raidant" : label;
      const value = cache.values[label] ?? cache.values[alt] ?? "-";
      $(row).find('.sheexcel-armor-value').text(value);
    });
  }

  async _locateStatsCells(sheetName) {
    const scan = await this._scanArea(sheetName, "A1", "Z60");
    const normalize = (value) => String(value || "").trim().replace(/:$/, "").toLowerCase();
    const skipLabels = new Set(["vitality", "health"]);

    const findHeaderRow = () => {
      for (let row = scan.startRow; row <= scan.endRow; row++) {
        for (let col = scan.startColumn; col <= scan.endColumn; col++) {
          if (normalize(scan.readCell(row, col)) === "stats") {
            return { row, col };
          }
        }
      }
      return null;
    };

    const statsRef = findHeaderRow();
    if (!statsRef) return null;

    const findBaseTot = (row) => {
      let base = null;
      let tot = null;
      for (let col = statsRef.col + 1; col <= scan.endColumn; col++) {
        const header = normalize(scan.readCell(row, col));
        if (header === "base") base = col;
        if (header === "tot" || header === "total") tot = col;
      }
      return { base, tot };
    };

    let headerRow = statsRef.row;
    let { base: baseCol, tot: totCol } = findBaseTot(headerRow);
    if (!baseCol && !totCol) {
      headerRow = Math.min(scan.endRow, statsRef.row + 1);
      ({ base: baseCol, tot: totCol } = findBaseTot(headerRow));
    }

    const valueCol = totCol || baseCol || (statsRef.col + 1);
    const startRow = headerRow + 1;
    const entries = [];
    const maxRow = Math.min(scan.endRow, startRow + 24);
    for (let row = startRow; row <= maxRow; row++) {
      const labelRaw = String(scan.readCell(row, statsRef.col) || "").trim();
      if (!labelRaw) continue;
      const normalized = normalize(labelRaw);
      if (skipLabels.has(normalized)) continue;
      const displayLabel = labelRaw.replace(/:$/, "").trim();
      const baseCell = baseCol ? this._buildSingleCell(baseCol, row) : null;
      const totCell = totCol ? this._buildSingleCell(totCol, row) : this._buildSingleCell(valueCol, row);
      entries.push({
        label: normalized,
        displayLabel,
        baseCell,
        totCell
      });
      if (normalized === "exhaustion") break;
    }

    if (!entries.length) return null;
    return { sheetName, entries };
  }

  async _getStatsCells(options = {}) {
    const { force = false } = options;
    const sheetId = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_ID);
    if (!sheetId) return null;

    if (this._statsCells?.sheetName) {
      const hasDisplay = Array.isArray(this._statsCells.entries)
        && this._statsCells.entries.every(entry => entry.displayLabel);
      if (hasDisplay) {
        return { sheetId, ...this._statsCells };
      }
    }

    const now = Date.now();
    if (!force && this._statsLocateAttempt && (now - this._statsLocateAttempt) < 60000) {
      return null;
    }
    this._statsLocateAttempt = now;

    const metadata = await this._fetchSheetMetadata();
    const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
    if (!sheets.length) return null;

    const coreSheet = sheets.find(s => /^core$/i.test(s.properties?.title || ""))
      || sheets.find(s => /core/i.test(s.properties?.title || ""));
    if (!coreSheet?.properties?.title) return null;

    const sheetName = coreSheet.properties.title;
    const located = await this._locateStatsCells(sheetName);
    if (!located) return null;

    this._statsCells = located;
    return { sheetId, ...located };
  }

  async _updateStatsBlock(html, options = {}) {
    const { force = false } = options;
    if (this._statsUpdateInFlight) return;
    this._statsUpdateInFlight = true;
    try {
      const cells = await this._getStatsCells({ force });
      if (!cells) return;

      const safeSheet = this._sanitizeSheetName(cells.sheetName);
      const ranges = [];
      cells.entries.forEach(entry => {
        if (entry.totCell) ranges.push(`${safeSheet}!${entry.totCell}`);
        if (entry.baseCell) ranges.push(`${safeSheet}!${entry.baseCell}`);
      });
      const response = await this._fetchValuesNoCache(cells.sheetId, ranges);
      const values = response.valueRanges || [];

      const normalizeRange = (range) => {
        const raw = String(range || "");
        const parts = raw.split("!");
        return parts.length > 1 ? parts[1].replace(/\$/g, "") : raw.replace(/\$/g, "");
      };

      const valueByCell = new Map();
      values.forEach((rangeEntry) => {
        const cellKey = normalizeRange(rangeEntry?.range);
        if (!cellKey) return;
        const raw = String(rangeEntry?.values?.[0]?.[0] ?? "").trim();
        valueByCell.set(cellKey, raw);
      });

      const valueMap = new Map();
      cells.entries.forEach((entry) => {
        const totRaw = entry.totCell ? (valueByCell.get(entry.totCell) ?? "") : "";
        const baseRaw = entry.baseCell ? (valueByCell.get(entry.baseCell) ?? "") : "";
        const raw = String(totRaw || baseRaw || "").trim();
        const numberMatch = raw.match(/[+-]?\d+(?:\.\d+)?/);
        valueMap.set(entry.label, numberMatch?.[0] ?? raw ?? "-");
      });

      const grid = html.find('.sheexcel-stats-grid');
      grid.attr("data-stats-sheet", cells.sheetName || "");
      grid.empty();
      const entriesCache = [];
      cells.entries.forEach((entry) => {
        const value = valueMap.get(entry.label) ?? "-";
        const label = entry.displayLabel || entry.label || "";
        const totCell = entry.totCell || "";
        const row = `
          <div class="sheexcel-stats-row" data-stats-label="${entry.label}" data-stats-cell="${totCell}">
            <span class="sheexcel-stats-label">${label}</span>
            <span class="sheexcel-stats-value">${value}</span>
          </div>
        `;
        grid.append(row);
        entriesCache.push({
          label: entry.label,
          displayLabel: label,
          value,
          totCell,
          baseCell: entry.baseCell || ""
        });
      });

      await this.actor.setFlag(MODULE_NAME, FLAGS.STATS_CACHE, {
        entries: entriesCache,
        updatedAt: Date.now(),
        sheetName: cells.sheetName || ""
      });
    } finally {
      this._statsUpdateInFlight = false;
    }
  }

  _startStatsPolling(html) {
    this._updateStatsBlock(html).catch(() => {});
  }

  _loadStatsFromFlags(html) {
    const cache = this.actor.getFlag(MODULE_NAME, FLAGS.STATS_CACHE);
    if (!cache?.entries?.length) return;

    const grid = html.find('.sheexcel-stats-grid');
    grid.attr("data-stats-sheet", cache.sheetName || "");
    grid.empty();
    cache.entries.forEach((entry) => {
      const label = entry.displayLabel || entry.label || "";
      const totCell = entry.totCell || "";
      const value = entry.value ?? "-";
      const row = `
        <div class="sheexcel-stats-row" data-stats-label="${entry.label}" data-stats-cell="${totCell}">
          <span class="sheexcel-stats-label">${label}</span>
          <span class="sheexcel-stats-value">${value}</span>
        </div>
      `;
      grid.append(row);
    });
  }

  async _getStatCells(labels, cacheKey) {
    const sheetId = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_ID);
    if (!sheetId) return null;

    const now = Date.now();
    const cooldownMs = 60000;
    const lastAttemptKey = `${cacheKey}AttemptAt`;

    if (this[cacheKey] && this[cacheKey].sheetName) {
      return { sheetId, sheetName: this[cacheKey].sheetName, ...this[cacheKey] };
    }

    if (this[lastAttemptKey] && (now - this[lastAttemptKey]) < cooldownMs) {
      return null;
    }

    this[lastAttemptKey] = now;

    const metadata = await this._fetchSheetMetadata();
    const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
    if (!sheets.length) return null;

    const coreSheet = sheets.find(s => /^core$/i.test(s.properties?.title || ""))
      || sheets.find(s => /core/i.test(s.properties?.title || ""));
    if (!coreSheet?.properties?.title) return null;

    const sheetName = coreSheet.properties.title;
    this[cacheKey] = await this._locateStatCells(sheetName, labels);
    if (!this[cacheKey]) return null;

    return {
      sheetId,
      sheetName,
      ...this[cacheKey]
    };
  }

  // Read only the BASE (max) value for a stat from the Google Sheet.
  async _getMaxFromSheet(labels, cacheKey) {
    const cells = await this._getStatCells(labels, cacheKey);
    if (!cells) return null;
    const safeSheet = this._sanitizeSheetName(cells.sheetName);
    const range = `${safeSheet}!${cells.baseCell}`;
    const response = await this._fetchValuesNoCache(cells.sheetId, [range]);
    const raw = response.valueRanges?.[0]?.values?.[0]?.[0] ?? "";
    const max = this._parseRawNumber(raw);
    if (!Number.isFinite(max) || max < 0) return null;
    return max;
  }

  // Apply current orb state (CSS vars + labels) from flags + cached max values.
  _applyPoeOrbValues(html) {
    const maxHp  = this._orbMaxHp  ?? this.actor.getFlag(MODULE_NAME, FLAGS.ORB_MAX_HP)  ?? 0;
    const maxVit = this._orbMaxVit ?? this.actor.getFlag(MODULE_NAME, FLAGS.ORB_MAX_VIT) ?? 0;
    const currentHp  = this.actor.getFlag(MODULE_NAME, FLAGS.ORB_CURRENT_HP)  ?? maxHp;
    const currentVit = this.actor.getFlag(MODULE_NAME, FLAGS.ORB_CURRENT_VIT) ?? (maxVit > 0 ? maxVit : 0);
    const customFrameImage = String(this.actor.getFlag(MODULE_NAME, FLAGS.ORB_FRAME_IMAGE) || "").trim();

    const lifeFill = maxHp  > 0 ? Math.min(1, Math.max(0, currentHp  / maxHp))  : 1;
    const vitFill  = maxVit > 0 ? Math.min(1, Math.max(0, currentVit / maxVit)) : 0;

    const orb = html.find('.sheexcel-poe-orb');
    if (!orb.length) return;
    const orbBorder = html.find('.sheexcel-poe-orb-border');

    orb[0].style.setProperty('--life-fill', lifeFill.toFixed(4));
    orb[0].style.setProperty('--vit-fill',  vitFill.toFixed(4));
    if (orbBorder.length) {
      orbBorder[0].style.backgroundImage = customFrameImage ? `url("${customFrameImage.replace(/"/g, '\\"')}")` : "";
    }

    html.find('.sheexcel-poe-life-value').text(
      maxHp > 0 ? `${Math.round(currentHp)}/${Math.round(maxHp)}` : '-'
    );
    html.find('.sheexcel-poe-shield-value').text(
      maxVit > 0 ? `${Math.round(currentVit)}/${Math.round(maxVit)}` : '0/0'
    );

    const shieldLayer = html.find('.sheexcel-poe-shield-layer');
    const shieldLabel = html.find('.sheexcel-poe-shield-label');
    if (maxVit <= 0) {
      shieldLayer.hide();
      shieldLabel.hide();
    } else {
      shieldLayer.show();
      shieldLabel.show();
    }

    orb.toggleClass('has-shield',    maxVit > 0 && currentVit > 0);
    orb.toggleClass('life-critical', maxHp  > 0 && (currentHp / maxHp) < 0.3);
  }

  // Fetch max values from sheet; cache them in flags so they survive reloads.
  async _updatePoeOrb(html) {
    if (this._orbUpdateInFlight) return;
    this._orbUpdateInFlight = true;
    try {
      const [maxHp, maxVit] = await Promise.all([
        this._getMaxFromSheet(["health", "hp"],    "_hpCells").catch(() => null),
        this._getMaxFromSheet(["vitality", "vit"], "_vitCells").catch(() => null)
      ]);

      let changed = false;
      if (maxHp !== null && maxHp !== this._orbMaxHp) {
        this._orbMaxHp = maxHp;
        await this.actor.setFlag(MODULE_NAME, FLAGS.ORB_MAX_HP, maxHp);
        changed = true;
      }
      if (maxVit !== null && maxVit !== this._orbMaxVit) {
        this._orbMaxVit = maxVit;
        await this.actor.setFlag(MODULE_NAME, FLAGS.ORB_MAX_VIT, maxVit);
        changed = true;
      }
      if (changed) this._applyPoeOrbValues(html);
    } finally {
      this._orbUpdateInFlight = false;
    }
  }

  _setupOrbClickHandlers(html) {
    html.find('.sheexcel-poe-orb').on('click.poeOrb', () => {
      this._openOrbEditDialog(html);
    });
  }

  _openOrbFrameDialog(html) {
    const currentFrameImage = String(this.actor.getFlag(MODULE_NAME, FLAGS.ORB_FRAME_IMAGE) || "").trim();

    const content = `
      <div style="padding:12px 4px 4px;display:flex;flex-direction:column;gap:10px;">
        <label style="color:#c7a86d;font-weight:700;font-family:'Fontin',serif;">Orb Frame Image</label>
        <input type="text" id="poe-orb-frame-input" value="${escapeHtml(currentFrameImage)}"
          placeholder="Leave empty for the default border"
          style="width:100%;background:#1a140f;border:1px solid #bfa05a;color:#f0dfb3;padding:5px 9px;border-radius:4px;font-size:0.95em;box-sizing:border-box;"/>
        <div style="display:flex;gap:8px;">
          <button type="button" id="poe-orb-frame-browse"
            style="flex:1;background:#241a10;border:1px solid #bfa05a;color:#f0dfb3;padding:5px 9px;border-radius:4px;font-family:'Fontin',serif;cursor:pointer;">Browse</button>
          <button type="button" id="poe-orb-frame-clear"
            style="flex:1;background:#1a0003;border:1px solid #8b3a1a;color:#ffb3a1;padding:5px 9px;border-radius:4px;font-family:'Fontin',serif;cursor:pointer;">Use Default</button>
        </div>
      </div>`;

    new Dialog({
      title: "Change Orb Frame",
      content,
      buttons: {
        ok: {
          label: "Confirm",
          callback: async (dialogHtml) => {
            const nextFrameImage = String(dialogHtml.find("#poe-orb-frame-input").val() || "").trim();
            if (nextFrameImage) {
              await this.actor.setFlag(MODULE_NAME, FLAGS.ORB_FRAME_IMAGE, nextFrameImage);
            } else {
              await this.actor.unsetFlag(MODULE_NAME, FLAGS.ORB_FRAME_IMAGE);
            }
            this._applyPoeOrbValues(html);
          }
        },
        cancel: { label: "Cancel" }
      },
      default: "ok",
      render: (dialogHtml) => {
        dialogHtml.find("#poe-orb-frame-browse").on("click", (event) => {
          event.preventDefault();
          new FilePicker({
            type: "image",
            current: dialogHtml.find("#poe-orb-frame-input").val() || currentFrameImage,
            callback: (path) => {
              dialogHtml.find("#poe-orb-frame-input").val(path);
            }
          }).render(true);
        });

        dialogHtml.find("#poe-orb-frame-clear").on("click", (event) => {
          event.preventDefault();
          dialogHtml.find("#poe-orb-frame-input").val("");
        });
      }
    }).render(true);
  }

  _openOrbEditDialog(html) {
    const maxHp  = this._orbMaxHp  ?? this.actor.getFlag(MODULE_NAME, FLAGS.ORB_MAX_HP)  ?? 0;
    const maxVit = this._orbMaxVit ?? this.actor.getFlag(MODULE_NAME, FLAGS.ORB_MAX_VIT) ?? 0;
    const currentHp  = this.actor.getFlag(MODULE_NAME, FLAGS.ORB_CURRENT_HP)  ?? maxHp;
    const currentVit = this.actor.getFlag(MODULE_NAME, FLAGS.ORB_CURRENT_VIT) ?? maxVit;

    const shieldRow = maxVit > 0 ? `
      <div class="form-group" style="display:flex;align-items:center;gap:10px;margin-top:8px;">
        <label style="color:#88d4ff;font-weight:700;min-width:54px;font-family:'Fontin',serif;">Vitality</label>
        <input type="number" id="poe-shield-input" value="${Math.round(currentVit)}" min="0" max="${Math.round(maxVit)}"
          style="flex:1;background:#000a1f;border:1px solid #1e7fd4;color:#d4eeff;padding:5px 9px;border-radius:4px;font-size:1em;"/>
        <span style="color:#555;font-size:0.85em;min-width:40px;">/ ${Math.round(maxVit)}</span>
      </div>` : '';

    const content = `
      <div style="padding:12px 4px 4px;">
        <div class="form-group" style="display:flex;align-items:center;gap:10px;">
          <label style="color:#ff9898;font-weight:700;min-width:54px;font-family:'Fontin',serif;">Health</label>
          <input type="number" id="poe-life-input" value="${Math.round(currentHp)}" min="0" max="${Math.round(maxHp)}"
            style="flex:1;background:#1a0003;border:1px solid #8b0000;color:#ffdddd;padding:5px 9px;border-radius:4px;font-size:1em;"/>
          <span style="color:#555;font-size:0.85em;min-width:40px;">/ ${Math.round(maxHp)}</span>
        </div>
        ${shieldRow}
      </div>`;

    new Dialog({
      title: "Set Health & Vitality",
      content,
      buttons: {
        frame: {
          label: "Change Frame",
          callback: () => {
            this._openOrbFrameDialog(html);
          }
        },
        ok: {
          label: "Confirm",
          callback: async (d) => {
            const rawHp  = parseInt(d.find("#poe-life-input").val());
            const safeHp = Math.min(maxHp,  Math.max(0, isNaN(rawHp)  ? currentHp  : rawHp));
            await this.actor.setFlag(MODULE_NAME, FLAGS.ORB_CURRENT_HP, safeHp);
            if (maxVit > 0) {
              const rawVit  = parseInt(d.find("#poe-shield-input").val());
              const safeVit = Math.min(maxVit, Math.max(0, isNaN(rawVit) ? currentVit : rawVit));
              await this.actor.setFlag(MODULE_NAME, FLAGS.ORB_CURRENT_VIT, safeVit);
            }
            this._applyPoeOrbValues(html);
          }
        },
        cancel: { label: "Cancel" }
      },
      default: "ok"
    }).render(true);
  }

  _startOrbPolling(html) {
    if (this._orbPoller) clearInterval(this._orbPoller);

    // Seed instance vars from cached flags so first render is instant.
    this._orbMaxHp  = this.actor.getFlag(MODULE_NAME, FLAGS.ORB_MAX_HP)  ?? null;
    this._orbMaxVit = this.actor.getFlag(MODULE_NAME, FLAGS.ORB_MAX_VIT) ?? null;

    this._applyPoeOrbValues(html);
    this._updatePoeOrb(html).catch(() => {});

    this._orbPoller = setInterval(() => {
      this._updatePoeOrb(html).catch(() => {});
    }, 5000);

    this._setupOrbClickHandlers(html);
  }

  _extractLabelValueEntries(scan, filterFn, options = {}) {
    const { requireNumericValue = false } = options;
    const entries = [];
    const excluded = new Set(["range", "accuracy", "critical", "damage", "special"]);
    const numberPattern = /^[-+]?\d+(\.\d+)?$/;

    for (let row = scan.startRow; row <= scan.endRow; row++) {
      for (let col = scan.startColumn; col <= scan.endColumn - 3; col++) {
        const rawLabel = scan.readCell(row, col);
        if (!rawLabel) continue;
        const hasColon = rawLabel.endsWith(":");
        const label = rawLabel.replace(/:$/, "").trim();
        if (!label) continue;
        if (!/[A-Za-z]/.test(label)) continue;
        if (numberPattern.test(label)) continue;

        const normalized = label.toLowerCase();
        if (excluded.has(normalized)) continue;

        const valueColumn = col + 3;
        const value = scan.readCell(row, valueColumn);
        if (!value) continue;
        if (value.endsWith(":")) continue;
        if (requireNumericValue && !numberPattern.test(value)) continue;

        if (!hasColon) {
          const stat1 = scan.readCell(row, col + 1);
          const stat2 = scan.readCell(row, col + 2);
          const looksLikeSkillRow = numberPattern.test(String(stat1).trim()) || numberPattern.test(String(stat2).trim());
          if (!looksLikeSkillRow) continue;
        }

        if (filterFn && !filterFn({ label, normalized, row, col, valueColumn, value })) {
          continue;
        }

        entries.push({ label, normalized, row, col, valueColumn, value });
      }
    }

    return entries;
  }

  _extractSpellBlocks(scan) {
    const labelMap = {
      "circle": "circle",
      "type": "spellType",
      "components": "components",
      "cast time": "castTime",
      "cost": "cost",
      "range": "range",
      "duration": "duration",
      "description": "description",
      "effect": "effect",
      "empower": "empower",
      "source": "source",
      "discipline": "discipline"
    };
    const multiFields = new Set(["components", "description", "effect", "empower"]);
    const normalizeLabel = (value) => String(value || "")
      .trim()
      .replace(/:$/, "")
      .toLowerCase();
    const rowHasLabel = (row) => {
      if (row < scan.startRow || row > scan.endRow) return false;
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const raw = scan.readCell(row, col);
        if (!raw || !raw.endsWith(":")) continue;
        const normalized = normalizeLabel(raw);
        if (labelMap[normalized]) return true;
      }
      return false;
    };
    const rowHasCircleLabel = (row) => {
      if (row < scan.startRow || row > scan.endRow) return false;
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const raw = scan.readCell(row, col);
        if (!raw || !raw.endsWith(":")) continue;
        if (normalizeLabel(raw) === "circle") return true;
      }
      return false;
    };
    const rowHasAnyValue = (row) => {
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        if (scan.readCell(row, col)) return true;
      }
      return false;
    };

    const nameBlocks = [];
    for (let row = scan.startRow; row <= scan.endRow; row++) {
      if (!rowHasCircleLabel(row + 1)) continue;
      if (row > scan.startRow && rowHasAnyValue(row - 1)) continue;
      if (rowHasLabel(row)) continue;

      let nameCol = null;
      let nameValue = "";
      let nonLabelCount = 0;
      let labelCount = 0;

      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const value = scan.readCell(row, col);
        if (!value) continue;
        if (String(value).trim().endsWith(":")) {
          labelCount += 1;
          continue;
        }
        if (!nameCol) {
          nameCol = col;
          nameValue = value;
        }
        nonLabelCount += 1;
      }

      if (!nameCol || !nameValue) continue;
      if (labelCount > 0) continue;
      if (nonLabelCount !== 1) continue;
      if (!/[A-Za-z]/.test(String(nameValue))) continue;

      nameBlocks.push({ row, col: nameCol, name: nameValue });
    }

    const spells = [];
    for (let i = 0; i < nameBlocks.length; i++) {
      const nameRow = nameBlocks[i].row;
      const nameCol = nameBlocks[i].col;
      const spellName = nameBlocks[i].name;
      const blockEnd = (nameBlocks[i + 1]?.row || (scan.endRow + 1)) - 1;
      if (!spellName) continue;

      const spell = {
        id: randomID(),
        type: "spells",
        sheet: "",
        keyword: spellName,
        spellName,
        spellNameCell: this._buildSingleCell(nameCol, nameRow),
        circle: "",
        spellType: "",
        components: "",
        castTime: "",
        cost: "",
        range: "",
        duration: "",
        description: "",
        effect: "",
        empower: "",
        source: "",
        discipline: "",
        value: "",
        subchecks: []
      };

      const multiData = {};
      let activeMulti = null;

      for (let row = nameRow + 1; row <= blockEnd; row++) {
        const labels = [];
        for (let col = scan.startColumn; col <= scan.endColumn; col++) {
          const raw = scan.readCell(row, col);
          if (!raw || !raw.endsWith(":")) continue;
          const normalized = normalizeLabel(raw);
          const field = labelMap[normalized];
          if (!field) continue;
          labels.push({ field, col });
        }

        if (labels.length) {
          activeMulti = null;
          const sorted = labels.slice().sort((a, b) => a.col - b.col);
          for (const label of labels) {
            const nextLabel = sorted.find(l => l.col > label.col);
            const endCol = nextLabel ? nextLabel.col - 1 : scan.endColumn;
            const values = [];
            let valueCol = null;
            for (let col = label.col + 1; col <= endCol; col++) {
              const cellValue = scan.readCell(row, col);
              if (!cellValue) continue;
              if (!valueCol) valueCol = col;
              values.push(cellValue);
            }
            if (!valueCol || !values.length) continue;
            const value = values.join(" ");
            if (multiFields.has(label.field)) {
              activeMulti = {
                field: label.field,
                col: valueCol,
                endCol,
                startRow: row,
                endRow: row,
                values: value ? [value] : []
              };
              if (!multiData[label.field]) multiData[label.field] = activeMulti;
            } else {
              spell[label.field] = value;
              const range = valueCol === endCol
                ? this._buildSingleCell(valueCol, row)
                : this._buildAbsoluteRange(valueCol, endCol, row);
              spell[`${label.field}Cell`] = range;
            }
          }
          continue;
        }

        if (activeMulti) {
          const rowValues = [];
          for (let col = activeMulti.col; col <= (activeMulti.endCol || activeMulti.col); col++) {
            const cellValue = scan.readCell(row, col);
            if (!cellValue) continue;
            rowValues.push(cellValue);
          }
          if (rowValues.length) {
            activeMulti.endRow = row;
            activeMulti.values.push(rowValues.join(" "));
          } else if (!rowHasAnyValue(row)) {
            activeMulti = null;
          }
        }
      }

      for (const [field, data] of Object.entries(multiData)) {
        if (!data.values.length) continue;
        const joiner = field === "components" ? ", " : "\n";
        spell[field] = data.values.join(joiner);
        const range = (data.startRow === data.endRow && (data.endCol ?? data.col) === data.col)
          ? this._buildSingleCell(data.col, data.startRow)
          : this._buildRectRange(data.col, data.startRow, data.endCol ?? data.col, data.endRow);
        spell[`${field}Cell`] = range;
      }

      spells.push(spell);
    }

    return spells;
  }

  _extractAbilityBlocks(scan) {
    console.log("Sheexcel | Abilities parser start", {
      startRow: scan.startRow,
      endRow: scan.endRow,
      startColumn: scan.startColumn,
      endColumn: scan.endColumn
    });

    const getRowCells = (row) => {
      const cells = [];
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const value = scan.readCell(row, col);
        if (!value) continue;
        cells.push({ col, value: String(value).trim() });
      }
      return cells;
    };

    const isHeaderRow = (cells) => {
      if (!cells.length) return false;
      const first = cells[0];
      const second = cells[1];
      if (!first || !second) return false;
      const codeLike = /^[A-Z]{2,4}$/.test(first.value);
      const hasTitle = /[A-Za-z]/.test(second.value);
      return codeLike && hasTitle;
    };

    const headers = [];
    for (let row = scan.startRow; row <= scan.endRow; row++) {
      const cells = getRowCells(row);
      if (!isHeaderRow(cells)) continue;
      headers.push({ row, cells });
    }

    console.log("Sheexcel | Abilities header detection", {
      count: headers.length,
      sample: headers.slice(0, 20).map(h => ({
        row: h.row,
        code: h.cells[0]?.value || "",
        titleParts: h.cells.slice(1).map(c => c.value)
      }))
    });

    const abilities = [];
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      const nextHeaderRow = headers[i + 1]?.row || (scan.endRow + 1);
      const code = header.cells[0]?.value || "";
      const titleCells = header.cells.slice(1);
      const abilityName = titleCells.map(c => c.value).join(" ").replace(/\s+/g, " ").trim();
      if (!abilityName) continue;

      const titleStartCol = titleCells[0]?.col || header.cells[0].col;
      const titleEndCol = titleCells[titleCells.length - 1]?.col || titleStartCol;
      const abilityNameCell = titleStartCol === titleEndCol
        ? this._buildSingleCell(titleStartCol, header.row)
        : this._buildAbsoluteRange(titleStartCol, titleEndCol, header.row);

      const contentLines = [];
      const contentStartRow = header.row + 1;
      const contentEndRow = nextHeaderRow - 1;

      for (let row = contentStartRow; row <= contentEndRow; row++) {
        const cells = getRowCells(row);
        if (!cells.length) continue;

        // Defensive: skip accidental header-like rows inside block.
        if (isHeaderRow(cells)) break;

        const minCol = cells[0].col;
        const indentLevel = Math.max(0, minCol - titleStartCol);
        const indent = "  ".repeat(Math.min(indentLevel, 6));
        const line = cells.map(c => c.value).join(" ").replace(/\s+/g, " ").trim();
        if (!line) continue;
        contentLines.push(`${indent}${line}`);
      }

      const effectText = contentLines.join("\n").trim();
      const ability = {
        id: randomID(),
        type: "abilities",
        sheet: "",
        keyword: abilityName,
        abilityName,
        abilityNameCell,
        category: code,
        cost: "",
        trigger: "",
        target: "",
        range: "",
        duration: "",
        description: "",
        effect: effectText,
        notes: "",
        value: "",
        subchecks: []
      };

      abilities.push(ability);
    }

    console.log("Sheexcel | Abilities parser end", {
      parsedAbilities: abilities.length,
      sample: abilities.slice(0, 10).map(a => ({
        name: a.abilityName,
        category: a.category,
        hasEffect: Boolean(a.effect)
      }))
    });

    return abilities;
  }

  async _onBulkAddChecks(event) {
    event.preventDefault();

    const sheetNames = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_NAMES) || [];
    const defaultSheet = sheetNames[0] || "Core";
    const options = sheetNames.map(name => `<option value="${name}">${name}</option>`).join("");

    const content = `
      <form class="sheexcel-bulk-checks-form" style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
        <label>Sheet</label>
        <select name="sheet">${options || `<option value="${defaultSheet}">${defaultSheet}</option>`}</select>

        <label>Top Left Cell</label>
        <input name="topLeftCell" type="text" value="A17" />

        <label>Bottom Right Cell</label>
        <input name="bottomRightCell" type="text" value="P60" />
      </form>
      <div style="margin-top:10px;padding:8px 10px;border:1px solid rgba(191,160,90,0.35);border-radius:6px;background:rgba(191,160,90,0.08);">
        <p style="margin:0 0 6px 0;font-weight:600;">Checks scanner</p>
        <p style="margin:2px 0;">Scans label + value pairs in this rectangle and creates checks, with consecutive rows grouped as subchecks.</p>
      </div>
    `;

    new Dialog({
      title: "Bulk Add Checks",
      content,
      buttons: {
        add: {
          label: "Add",
          callback: async (html) => {
            try {
              const form = html[0].querySelector(".sheexcel-bulk-checks-form");
              const data = new FormData(form);
              const sheet = String(data.get("sheet") || defaultSheet);
              const topLeftCell = String(data.get("topLeftCell") || "A17");
              const bottomRightCell = String(data.get("bottomRightCell") || "P60");

              const scan = await this._scanArea(sheet, topLeftCell, bottomRightCell);
              const entries = this._extractLabelValueEntries(
                scan,
                ({ normalized }) => !normalized.includes("save"),
                { requireNumericValue: true }
              );
              if (!entries.length) throw new Error("No checks detected in selected range.");

              const existingRefs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
              await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, existingRefs);

              const grouped = new Map();
              for (const entry of entries) {
                const key = entry.col;
                if (!grouped.has(key)) grouped.set(key, []);
                grouped.get(key).push(entry);
              }

              const additions = [];
              let subcheckCount = 0;

              for (const [, list] of grouped) {
                list.sort((a, b) => a.row - b.row);
                let parent = null;
                let lastRow = -999;

                for (const entry of list) {
                  const cell = this._buildSingleCell(entry.valueColumn, entry.row);
                  const isNextRow = parent && (entry.row - lastRow) === 1;
                  if (!isNextRow) {
                    parent = {
                      id: randomID(),
                      cell,
                      sheet,
                      keyword: entry.label,
                      type: "checks",
                      value: "",
                      attackNameCell: "",
                      critRangeCell: "",
                      damageCell: "",
                      subchecks: []
                    };
                    additions.push(parent);
                  } else {
                    parent.subchecks.push({
                      id: randomID(),
                      cell,
                      sheet,
                      keyword: entry.label,
                      type: "checks",
                      value: "",
                      attackNameCell: "",
                      critRangeCell: "",
                      damageCell: "",
                      subchecks: []
                    });
                    subcheckCount += 1;
                  }

                  lastRow = entry.row;
                }
              }

              if (!additions.length) throw new Error("No checks generated from detected entries.");

              await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, [...existingRefs, ...additions]);
              ui.notifications.info(`Added ${additions.length} checks and ${subcheckCount} subchecks. Use Undo Bulk Add to revert.`);
              this.render(false);
            } catch (error) {
              ui.notifications.error(`Bulk checks failed: ${error.message}`);
            }
          }
        },
        cancel: { label: "Cancel" }
      },
      default: "add"
    }).render(true);
  }

  async _onUpdateSkillsFromSheet(event) {
    event.preventDefault();

    try {
      this._invalidateApiCache();
      const metadata = await this._fetchSheetMetadata();
      const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
      if (!sheets.length) throw new Error("No sheet metadata found.");

      const coreSheet = sheets.find(s => /^core$/i.test(s.properties?.title || ""))
        || sheets.find(s => /core/i.test(s.properties?.title || ""));
      if (!coreSheet?.properties?.title) throw new Error("Core sheet tab not found.");

      const sheetName = coreSheet.properties.title;
      const rowCount = Math.max(2, Math.min(400, Number(coreSheet.properties?.gridProperties?.rowCount) || 200));
      const columnCount = Math.max(2, Math.min(26, Number(coreSheet.properties?.gridProperties?.columnCount) || 16));
      const searchBottomRight = `${this._columnNumberToLetters(columnCount)}${rowCount}`;

      const searchScan = await this._scanArea(sheetName, "A1", searchBottomRight);
      const normalizeLabel = (value) => String(value || "").trim().replace(/:$/, "").toLowerCase();
      const numberLike = /^[-+]?\d+(\.\d+)?$/;

      let athletics = null;
      let languages = null;
      let initiativeRef = null;
      for (let row = searchScan.startRow; row <= searchScan.endRow; row++) {
        for (let col = searchScan.startColumn; col <= searchScan.endColumn; col++) {
          const raw = searchScan.readCell(row, col);
          if (!raw) continue;
          const normalized = normalizeLabel(raw);
          if (!athletics && normalized === "athletics") {
            athletics = { row, col };
          }
          if (normalized === "languages" || normalized === "language") {
            if (!languages || row > languages.row || (row === languages.row && col > languages.col)) {
              languages = { row, col };
            }
          }
          if (!initiativeRef && normalized === "initiative") {
            let valueCol = null;
            let value = "";
            for (let c = col + 1; c <= searchScan.endColumn; c++) {
              const candidate = String(searchScan.readCell(row, c) || "").trim();
              if (!candidate) continue;
              if (!numberLike.test(candidate)) continue;
              valueCol = c;
              value = candidate;
              break;
            }
            if (valueCol) {
              initiativeRef = {
                id: randomID(),
                cell: this._buildSingleCell(valueCol, row),
                sheet: sheetName,
                keyword: "Initiative",
                type: "checks",
                value,
                attackNameCell: "",
                critRangeCell: "",
                damageCell: "",
                subchecks: []
              };
            }
          }
        }
      }

      if (!athletics) throw new Error("Could not find 'Athletics' on Core sheet.");
      if (!languages) throw new Error("Could not find 'Languages' on Core sheet.");
      if (languages.row < athletics.row || languages.col < athletics.col) {
        throw new Error("Detected Languages appears above/left of Athletics. Check Core layout.");
      }

      const topLeft = this._buildSingleCell(athletics.col, athletics.row);
      const bottomRightCol = Math.min(searchScan.endColumn, languages.col + 3);
      const hasDataInWindow = (row, startCol, endCol) => {
        for (let col = startCol; col <= endCol; col++) {
          if (searchScan.readCell(row, col)) return true;
        }
        return false;
      };

      let bottomRightRow = languages.row;
      let seenAnyAfterLanguage = false;
      for (let row = languages.row + 1; row <= searchScan.endRow; row++) {
        if (hasDataInWindow(row, athletics.col, bottomRightCol)) {
          bottomRightRow = row;
          seenAnyAfterLanguage = true;
          continue;
        }
        if (seenAnyAfterLanguage) break;
      }

      const bottomRight = this._buildSingleCell(bottomRightCol, bottomRightRow);
      const scan = await this._scanArea(sheetName, topLeft, bottomRight);
      const notes = await this._fetchCellNotes(sheetName, topLeft, bottomRight);

      const entries = this._extractLabelValueEntries(
        scan,
        ({ normalized }) => !normalized.includes("save"),
        { requireNumericValue: true }
      );
      if (!entries.length) throw new Error("No checks detected in Core skills area.");

      const grouped = new Map();
      for (const entry of entries) {
        const key = entry.col;
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key).push(entry);
      }

      const additions = [];
      let subcheckCount = 0;

      for (const [, list] of grouped) {
        list.sort((a, b) => a.row - b.row);
        let parent = null;
        let lastRow = -999;

        for (const entry of list) {
          const cell = this._buildSingleCell(entry.valueColumn, entry.row);
          const isNextRow = parent && (entry.row - lastRow) === 1;
          if (!isNextRow) {
            const noteKey = `${entry.row},${entry.col}`;
            const leftValueKey = `${entry.row},${entry.valueColumn - 1}`;
            const comment = notes.get(leftValueKey) || notes.get(noteKey) || "";
            parent = {
              id: randomID(),
              cell,
              sheet: sheetName,
              keyword: entry.label,
              type: "checks",
              value: entry.value,
              comment,
              attackNameCell: "",
              critRangeCell: "",
              damageCell: "",
              subchecks: []
            };
            additions.push(parent);
          } else {
            const noteKey = `${entry.row},${entry.col}`;
            const leftValueKey = `${entry.row},${entry.valueColumn - 1}`;
            const comment = notes.get(leftValueKey) || notes.get(noteKey) || "";
            parent.subchecks.push({
              id: randomID(),
              cell,
              sheet: sheetName,
              keyword: entry.label,
              type: "checks",
              value: entry.value,
              comment,
              attackNameCell: "",
              critRangeCell: "",
              damageCell: "",
              subchecks: []
            });
            subcheckCount += 1;
          }
          lastRow = entry.row;
        }
      }

      if (!additions.length) throw new Error("No checks generated from detected entries.");

      if (initiativeRef) {
        const hasInitiative = additions.some(ref => String(ref.keyword || "").toLowerCase() === "initiative");
        if (!hasInitiative) {
          additions.unshift(initiativeRef);
        }
      }

      const refs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
      await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, refs);
      const nonChecks = refs.filter(ref => ref?.type !== "checks");
      await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, [...nonChecks, ...additions]);

      ui.notifications.info(`Updated skills from Core (${topLeft} → ${bottomRight}): ${additions.length} checks and ${subcheckCount} subchecks.`);
      this.render(false);
    } catch (error) {
      ui.notifications.error(`Update skills failed: ${error.message}`);
    }
  }

  async _onUpdateSavesFromSheet(event) {
    event.preventDefault();

    try {
      this._invalidateApiCache();
      const metadata = await this._fetchSheetMetadata();
      const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
      if (!sheets.length) throw new Error("No sheet metadata found.");

      const coreSheet = sheets.find(s => /^core$/i.test(s.properties?.title || ""))
        || sheets.find(s => /core/i.test(s.properties?.title || ""));
      if (!coreSheet?.properties?.title) throw new Error("Core sheet tab not found.");

      const sheetName = coreSheet.properties.title;
      const rowCount = Math.max(2, Math.min(400, Number(coreSheet.properties?.gridProperties?.rowCount) || 200));
      const columnCount = Math.max(2, Math.min(26, Number(coreSheet.properties?.gridProperties?.columnCount) || 16));
      const searchBottomRight = `${this._columnNumberToLetters(columnCount)}${rowCount}`;
      const searchScan = await this._scanArea(sheetName, "A1", searchBottomRight);

      const normalizeLabel = (value) => String(value || "").trim().replace(/:$/, "").toLowerCase();
      const saveOrder = ["Strength", "Dexterity", "Constitution", "Intelligence", "Wisdom", "Charisma"];
      const abilityMap = {
        strength: "Strength",
        dexterity: "Dexterity",
        constitution: "Constitution",
        intelligence: "Intelligence",
        wisdom: "Wisdom",
        charisma: "Charisma"
      };

      const attributesAnchors = [];
      const saveAnchors = [];
      for (let row = searchScan.startRow; row <= searchScan.endRow; row++) {
        for (let col = searchScan.startColumn; col <= searchScan.endColumn; col++) {
          const raw = searchScan.readCell(row, col);
          if (!raw) continue;
          const normalized = normalizeLabel(raw);
          if (normalized === "attributes") attributesAnchors.push({ row, col });
          if (normalized === "save") saveAnchors.push({ row, col });
        }
      }

      if (!attributesAnchors.length) throw new Error("Could not find 'Attributes' on Core sheet.");
      if (!saveAnchors.length) throw new Error("Could not find 'Save' on Core sheet.");

      let bestPair = null;
      for (const attr of attributesAnchors) {
        for (const save of saveAnchors) {
          if (save.row !== attr.row) continue;
          if (save.col <= attr.col) continue;
          const distance = save.col - attr.col;
          if (!bestPair || distance < bestPair.distance) {
            bestPair = { attr, save, distance };
          }
        }
      }
      if (!bestPair) throw new Error("Could not match 'Attributes' and 'Save' on the same row.");

      const startRow = bestPair.attr.row + 1;
      const endRow = searchScan.endRow;
      const abilityCol = bestPair.attr.col;
      const saveCol = bestPair.save.col;

      const foundByAbility = new Map();
      for (let row = startRow; row <= endRow; row++) {
        const abilityRaw = searchScan.readCell(row, abilityCol);
        const valueRaw = String(searchScan.readCell(row, saveCol) || "").trim();
        if (!abilityRaw && !valueRaw) {
          if (foundByAbility.size > 0) break;
          continue;
        }

        const abilityKey = normalizeLabel(abilityRaw);
        const canonical = abilityMap[abilityKey];
        if (!canonical) continue;
        if (!valueRaw) continue;
        if (!foundByAbility.has(canonical)) {
          foundByAbility.set(canonical, { row, value: valueRaw });
        }
      }

      const additions = [];
      for (const ability of saveOrder) {
        const hit = foundByAbility.get(ability);
        if (!hit) continue;
        additions.push({
          id: randomID(),
          cell: this._buildSingleCell(saveCol, hit.row),
          sheet: sheetName,
          keyword: ability,
          type: "saves",
          value: hit.value,
          attackNameCell: "",
          critRangeCell: "",
          damageCell: "",
          subchecks: []
        });
      }

      if (!additions.length) throw new Error("No saves detected under 'Attributes'/'Save' columns.");

      const refs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
      await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, refs);
      const nonSaves = refs.filter(ref => ref?.type !== "saves");
      await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, [...nonSaves, ...additions]);

      ui.notifications.info(`Updated saves from Core (${additions.length} found): ${saveOrder.join(", ")}.`);
      this.render(false);
    } catch (error) {
      ui.notifications.error(`Update saves failed: ${error.message}`);
    }
  }

  _extractCoreAbilityMods(scan) {
    const normalize = (value) => String(value || "").trim().replace(/:$/, "").toLowerCase();
    const labelMap = new Map([
      ["strength", "STR"],
      ["dexterity", "DEX"],
      ["constitution", "CON"],
      ["intelligence", "INT"],
      ["wisdom", "WIS"],
      ["charisma", "CHA"]
    ]);
    let modCol = null;
    let modRow = null;

    for (let row = scan.startRow; row <= scan.endRow; row++) {
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const cell = normalize(scan.readCell(row, col));
        if (cell === "mod") {
          modCol = col;
          modRow = row;
          break;
        }
      }
      if (modCol !== null) break;
    }

    const mods = {};
    if (modCol === null || modRow === null) return mods;

    const maxRow = Math.min(scan.endRow, modRow + 10);
    for (let row = modRow + 1; row <= maxRow; row++) {
      const labelLeft = normalize(scan.readCell(row, modCol - 1));
      const labelFar = normalize(scan.readCell(row, modCol - 2));
      const key = labelMap.get(labelLeft) || labelMap.get(labelFar);
      if (!key) continue;
      const raw = String(scan.readCell(row, modCol) || "").trim();
      const match = raw.match(/[-+]?\d+/);
      if (!match) continue;
      mods[key] = Number(match[0]);
    }

    return mods;
  }

  _extractCoreProficiency(scan) {
    const normalize = (value) => String(value || "")
      .trim()
      .toLowerCase()
      .replace(/[:\s]+/g, "");
    const findNumberInRow = (row, startCol, endCol) => {
      for (let vc = startCol; vc <= endCol; vc++) {
        const raw = String(scan.readCell(row, vc) || "").trim();
        if (!raw) continue;
        const match = raw.match(/[-+]?\d+/);
        if (match) return Number(match[0]);
      }
      return null;
    };

    for (let row = scan.startRow; row <= scan.endRow; row++) {
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const cell = normalize(scan.readCell(row, col));
        if (!cell.includes("proficiency") && !cell.includes("proficency") && !cell.includes("profiency")) continue;
        // First try the same row to the right of the label.
        const sameRow = findNumberInRow(row, col + 1, scan.endColumn);
        if (sameRow !== null) return sameRow;
        // If not found, try the next row (common in merged headers).
        const nextRow = row + 1 <= scan.endRow
          ? findNumberInRow(row + 1, col, scan.endColumn)
          : null;
        if (nextRow !== null) return nextRow;
      }
    }
    return 0;
  }

  _parseGearAbilityToken(text) {
    const match = String(text || "").match(/\(([^)]+)\)/);
    if (!match) return null;
    const token = match[1].trim().toLowerCase();
    if (token.startsWith("str")) return "STR";
    if (token.startsWith("dex")) return "DEX";
    if (token.startsWith("con")) return "CON";
    if (token.startsWith("int")) return "INT";
    if (token.startsWith("wis")) return "WIS";
    if (token.startsWith("cha")) return "CHA";
    return null;
  }

  _parseGearAccuracy(accuracy, mods, proficiency) {
    const text = String(accuracy || "").trim();
    if (!text) return null;
    const baseMatch = text.match(/[-+]?\d+/);
    const base = baseMatch ? Number(baseMatch[0]) : 0;
    const abilityKey = this._parseGearAbilityToken(text);
    const abilityMod = abilityKey && Number.isFinite(mods[abilityKey]) ? mods[abilityKey] : 0;
    const prof = Number.isFinite(proficiency) ? proficiency : 0;
    return {
      total: base + abilityMod + prof,
      base,
      abilityKey,
      abilityMod,
      proficiency: prof
    };
  }

  _parseGearCrit(crit) {
    const text = String(crit || "").trim();
    const match = text.match(/\d+/);
    return match ? Number(match[0]) : 20;
  }

  _parseGearDamage(damage, mods) {
    const text = String(damage || "").trim();
    if (!text) return "";
    const segment = text.split(/\band\b/i)[0];
    const diceMatch = segment.match(/\d+\s*d\s*\d+/i);
    if (!diceMatch) return "";
    const dice = diceMatch[0].replace(/\s+/g, "").toLowerCase();
    const abilityKey = this._parseGearAbilityToken(segment);
    const abilityMod = abilityKey && Number.isFinite(mods[abilityKey]) ? mods[abilityKey] : 0;
    const modTerm = abilityMod === 0 ? "" : abilityMod > 0 ? `+${abilityMod}` : `${abilityMod}`;
    return `${dice}${modTerm}`;
  }

  _formatSignedNumber(value) {
    if (!Number.isFinite(value)) return "0";
    return value >= 0 ? `+${value}` : String(value);
  }

  _buildAccuracyBreakdown(baseValue, abilityKey, abilityValue, profValue, accAdj, totalValue) {
    const abilityPart = abilityKey ? `${abilityKey} ${this._formatSignedNumber(abilityValue)}` : "";
    const parts = [
      this._formatSignedNumber(baseValue),
      abilityPart ? `+ ${abilityPart}` : "",
      `+ Prof ${this._formatSignedNumber(profValue)}`,
      accAdj ? `+ AccAdj ${this._formatSignedNumber(accAdj)}` : "",
      `= ${totalValue}`
    ];
    return parts.filter(Boolean).join(" ");
  }

  _parseAbilityKeyFromText(text) {
    const upper = String(text || "").toUpperCase();
    const plusMatch = upper.match(/\+\s*(STR|DEX|CON|INT|WIS|CHA)/);
    if (plusMatch) return plusMatch[1];
    return this._parseGearAbilityToken(text);
  }

  _parseAccuracyAdjustment(text) {
    const match = String(text || "").match(/([+-]\d+)\s*Accuracy/i);
    return match ? Number(match[1]) : null;
  }

  _parseDamageFormulaFromText(text, mods) {
    const diceMatch = String(text || "").match(/\d+\s*d\s*\d+/i);
    if (!diceMatch) return "";
    const dice = diceMatch[0].replace(/\s+/g, "").toLowerCase();
    const abilityKey = this._parseAbilityKeyFromText(text);
    const abilityMod = abilityKey && Number.isFinite(mods[abilityKey]) ? mods[abilityKey] : 0;
    const modTerm = abilityMod === 0 ? "" : abilityMod > 0 ? `+${abilityMod}` : `${abilityMod}`;
    return `${dice}${modTerm}`;
  }

  _extractDamageDetailFromText(text) {
    const match = String(text || "").match(/\d+\s*d\s*\d+[^.]*/i);
    return match ? match[0].trim() : "";
  }

  _extractDamageTypeFromText(text) {
    const source = String(text || "").toLowerCase();
    if (!source) return "";

    const knownTypes = [
      ["acid", "Acid"],
      ["blunt", "Bludgeoning"],
      ["bludgeoning", "Bludgeoning"],
      ["cold", "Cold"],
      ["fire", "Fire"],
      ["force", "Force"],
      ["lightning", "Lightning"],
      ["necrotic", "Necrotic"],
      ["pierce", "Piercing"],
      ["piercing", "Piercing"],
      ["poison", "Poison"],
      ["psychic", "Psychic"],
      ["radiant", "Radiant"],
      ["slash", "Slashing"],
      ["slashing", "Slashing"],
      ["thunder", "Thunder"],
      ["vitality", "Vitality"]
    ];

    const found = [];
    const seen = new Set();
    for (const [token, label] of knownTypes) {
      if (new RegExp(`\\b${token}\\b`, "i").test(source)) {
        if (seen.has(label)) continue;
        seen.add(label);
        found.push(label);
      }
    }
    if (found.length) return found.join(" / ");

    return "";
  }

  _extractDamagePartsFromText(text, mods = null) {
    const raw = String(text || "").trim();
    if (!raw) return [];

    const segments = raw
      .split(/\band\b|\//i)
      .map(part => part.trim())
      .filter(Boolean);

    const parts = [];
    const seen = new Set();
    for (const segment of segments) {
      const formula = mods
        ? (this._parseDamageFormulaFromText(segment, mods) || this._parseRawDamageFormula(segment))
        : this._parseRawDamageFormula(segment);
      if (!formula) continue;

      const type = this._extractDamageTypeFromText(segment);
      const key = `${formula}::${(type||"").toLowerCase()}`;
      if (seen.has(key)) continue;
      seen.add(key);
      parts.push({ formula, type, detail: segment });
    }

    if (!parts.length) {
      const fallbackFormula = mods
        ? (this._parseDamageFormulaFromText(raw, mods) || this._parseRawDamageFormula(raw))
        : this._parseRawDamageFormula(raw);
      if (fallbackFormula) {
        parts.push({
          formula: fallbackFormula,
          type: this._extractDamageTypeFromText(raw),
          detail: raw
        });
      }
    }

    return parts;
  }

  _parseRawNumber(text) {
    const match = String(text || "").match(/[-+]?\d+/);
    return match ? Number(match[0]) : null;
  }

  _parseRawDamageFormula(text) {
    const diceMatch = String(text || "").match(/\d+\s*d\s*\d+/i);
    if (!diceMatch) return "";
    const dice = diceMatch[0].replace(/\s+/g, "").toLowerCase();
    const tail = String(text || "").slice(diceMatch.index + diceMatch[0].length);
    const modMatch = tail.match(/[-+]?\s*\d+/);
    if (!modMatch) return dice;
    const mod = modMatch[0].replace(/\s+/g, "");
    return `${dice}${mod.startsWith("+") || mod.startsWith("-") ? mod : `+${mod}`}`;
  }

  _extractCoreAttackBlocks(scan) {
    const normalize = (value) => String(value || "").trim().replace(/:$/, "").toLowerCase();
    const labels = new Set(["range", "accuracy", "critical", "damage", "special"]);
    const attacks = [];
    const seen = new Set();

    const findValueRight = (row, col) => {
      for (let c = col + 1; c <= scan.endColumn; c++) {
        const raw = String(scan.readCell(row, c) || "").trim();
        if (!raw) continue;
        return { value: raw, col: c };
      }
      return null;
    };

    const collectValuesRight = (row, col) => {
      const values = [];
      let lastCol = col;
      let started = false;
      let emptyRun = 0;
      const maxGapAfterStart = 1;

      for (let c = col + 1; c <= scan.endColumn; c++) {
        const raw = String(scan.readCell(row, c) || "").trim();
        if (!raw) {
          if (started) {
            emptyRun += 1;
            if (emptyRun > maxGapAfterStart) break;
          }
          continue;
        }

        started = true;
        emptyRun = 0;
        values.push(raw);
        lastCol = c;
      }
      return { values, lastCol };
    };

    const findHeaderAbove = (startRow, col) => {
      const minRow = Math.max(scan.startRow, startRow - 20);
      const maxOffset = 12;
      const isHeader = (raw) => {
        if (!raw) return false;
        const normalized = normalize(raw);
        if (labels.has(normalized)) return false;
        if (/^[-+]?\d+(\.\d+)?$/.test(normalized)) return false;
        return true;
      };

      for (let r = startRow; r >= minRow; r--) {
        const sameCol = String(scan.readCell(r, col) || "").trim();
        if (isHeader(sameCol)) return { text: sameCol, row: r, col };

        for (let offset = 1; offset <= maxOffset; offset++) {
          const leftCol = col - offset;
          const rightCol = col + offset;

          if (leftCol >= scan.startColumn) {
            const left = String(scan.readCell(r, leftCol) || "").trim();
            if (isHeader(left)) return { text: left, row: r, col: leftCol };
          }

          if (rightCol <= scan.endColumn) {
            const right = String(scan.readCell(r, rightCol) || "").trim();
            if (isHeader(right)) return { text: right, row: r, col: rightCol };
          }
        }
      }

      return null;
    };

    for (let row = scan.startRow; row <= scan.endRow; row++) {
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const cell = normalize(scan.readCell(row, col));
        if (cell !== "range") continue;
        const rangeValue = findValueRight(row, col);
        if (!rangeValue) continue;

        const valueCol = rangeValue.col;
        const header = findHeaderAbove(row - 1, valueCol);
        const attackName = header?.text || `Attack ${row}`;
        let maxValueCol = valueCol;

        let accuracy = "";
        let critical = "";
        let damage = "";
        let damageRows = [];
        let special = "";
        let lastLabel = "";

        let endRow = Math.min(scan.endRow, row + 20);
        for (let r = row + 1; r <= endRow; r++) {
          const nextCell = normalize(scan.readCell(r, col));
          if (nextCell === "range") {
            endRow = r - 1;
            break;
          }
        }
        for (let r = row; r <= endRow; r++) {
          const label = normalize(scan.readCell(r, col));
          if (labels.has(label)) {
            lastLabel = label;
            const value = findValueRight(r, col);
            if (!value) continue;
            if (value.col > maxValueCol) maxValueCol = value.col;
            if (label === "accuracy") accuracy = value.value;
            if (label === "critical") critical = value.value;
            if (label === "damage") {
              const rowValues = collectValuesRight(r, col);
              if (rowValues.lastCol > maxValueCol) maxValueCol = rowValues.lastCol;
              damage = rowValues.values[0] || value.value;
              damageRows = rowValues.values.length ? [rowValues.values.join(" ")] : [];
            }
            if (label === "special") special = value.value;
            continue;
          }

          if (!label && lastLabel === "damage") {
            const rowValues = collectValuesRight(r, col);
            const value = rowValues.values[0] || "";
            if (rowValues.lastCol > maxValueCol) maxValueCol = rowValues.lastCol;
            if (value) {
              damage = damage ? `${damage} / ${value}` : value;
              damageRows.push(rowValues.values.join(" "));
            }
          }
        }

        const key = `${attackName}|${row}|${valueCol}`;
        if (seen.has(key)) continue;
        seen.add(key);

        const blockStartRow = header?.row ? header.row : row;
        const blockStartCol = Math.min(col, header?.col ?? col);
        const blockEndCol = Math.max(maxValueCol, blockStartCol);

        attacks.push({
          attackName,
          range: rangeValue.value,
          accuracy,
          critical,
          damage,
          damageRows,
          special,
          startRow: blockStartRow,
          endRow,
          startCol: blockStartCol,
          endCol: blockEndCol
        });
      }
    }

    return attacks;
  }

  _extractAbilityAttacksFromGear(gear, mods, accuracy, critRange, sheetName) {
    const abilitiesText = String(gear.abilities || "");
    if (!abilitiesText) return [];

    const clean = abilitiesText.replace(/<[^>]+>/g, "");
    const lines = clean
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean);

    const attacks = [];
    const seen = new Set();
    let currentLabel = "";
    let pendingAccAdj = null;

    for (const rawLine of lines) {
      const line = rawLine.replace(/\s+/g, " ").trim();
      if (!line) continue;

      let label = "";
      const labelMatch = line.match(/^([^:]+):\s*(.*)$/);
      if (labelMatch) {
        label = labelMatch[1].trim();
        currentLabel = label;
        const adjFromLabel = this._parseAccuracyAdjustment(line);
        if (adjFromLabel !== null) pendingAccAdj = adjFromLabel;
      }

      const hasDice = /\d+\s*d\s*\d+/i.test(line);
      const adjInLine = this._parseAccuracyAdjustment(line);
      if (adjInLine !== null) pendingAccAdj = adjInLine;
      if (!hasDice) continue;

      const abilityLabel = label || currentLabel || "Ability";
      if (!/thrust/i.test(abilityLabel)) {
        continue;
      }
      const damageParts = this._extractDamagePartsFromText(line, mods);
      if (!damageParts.length) continue;
      const damageFormula = damageParts[0].formula;
      const detail = damageParts.map(part => part.detail).join(" / ");
      const damageType = damageParts.map(part => part.type).filter(Boolean).join(" / ");
      const accAdj = pendingAccAdj || 0;

      const baseTotal = Number.isFinite(accuracy?.total) ? accuracy.total : 0;
      const totalValue = baseTotal + accAdj;
      const baseValue = Number.isFinite(accuracy?.base) ? accuracy.base : 0;
      const abilityKey = accuracy?.abilityKey || "";
      const abilityValue = Number.isFinite(accuracy?.abilityMod) ? accuracy.abilityMod : 0;
      const profValue = Number.isFinite(accuracy?.proficiency) ? accuracy.proficiency : 0;
      const accuracyBreakdown = this._buildAccuracyBreakdown(
        baseValue,
        abilityKey,
        abilityValue,
        profValue,
        accAdj,
        totalValue
      );

      const attackName = `${gear.gearName} - ${abilityLabel}`;
      const key = `${attackName}|${damageFormula}|${totalValue}|${detail}`;
      if (seen.has(key)) continue;
      seen.add(key);

      attacks.push({
        id: randomID(),
        type: "attacks",
        sheet: sheetName,
        keyword: gear.gearName,
        attackName,
        attackDetail: detail,
        damageType,
        damageParts,
        value: totalValue,
        accuracyBase: baseValue,
        accuracyAbility: abilityKey,
        accuracyAbilityMod: abilityValue,
        accuracyProficiency: profValue,
        accuracyAdjustment: accAdj,
        accuracyBreakdown,
        critRange: Number.isFinite(critRange) ? critRange : 20,
        damage: damageFormula,
        attackNameCell: "",
        critRangeCell: "",
        damageCell: ""
      });

      pendingAccAdj = null;
    }

    return attacks;
  }

  async _onUpdateAttacksFromSheet(event) {
    event.preventDefault();

    try {
      this._invalidateApiCache();
      const metadata = await this._fetchSheetMetadata();
      const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
      if (!sheets.length) throw new Error("No sheet metadata found.");

      const coreSheet = sheets.find(s => /^core$/i.test(s.properties?.title || ""))
        || sheets.find(s => /core/i.test(s.properties?.title || ""));
      if (!coreSheet?.properties?.title) throw new Error("Core sheet tab not found.");

      const coreRowCount = Math.max(2, Number(coreSheet.properties?.gridProperties?.rowCount) || 400);
      const coreColumnCount = Math.max(2, Number(coreSheet.properties?.gridProperties?.columnCount) || 40);
      const coreBottomRight = `${this._columnNumberToLetters(coreColumnCount)}${coreRowCount}`;

      const coreScan = await this._scanArea(coreSheet.properties.title, "A1", coreBottomRight);
      const coreNotes = await this._fetchCellNotes(coreSheet.properties.title, "A1", coreBottomRight);
      const coreAttacks = this._extractCoreAttackBlocks(coreScan);

      const collectNotesForBlock = (block) => {
        const startRow = Math.max(coreScan.startRow, Number(block?.startRow) || 0);
        const endRow = Math.min(coreScan.endRow, Number(block?.endRow) || 0);
        const startCol = Math.max(coreScan.startColumn, Number(block?.startCol) || 0);
        const endCol = Math.min(coreScan.endColumn, Number(block?.endCol) || 0);
        if (!startRow || !endRow || !startCol || !endCol) return "";

        const seen = new Set();
        const lines = [];
        for (let r = startRow; r <= endRow; r++) {
          for (let c = startCol; c <= endCol; c++) {
            const note = coreNotes.get(`${r},${c}`);
            if (!note) continue;
            if (seen.has(note)) continue;
            seen.add(note);
            lines.push(note);
          }
        }

        return lines.join("\n");
      };

      const attacks = coreAttacks
        .map(attack => {
          const accuracyValue = this._parseRawNumber(attack.accuracy) ?? 0;
          const critRange = this._parseRawNumber(attack.critical) ?? 20;
          const detail = attack.damage ? String(attack.damage).trim() : "";
          const damageSource = Array.isArray(attack.damageRows) && attack.damageRows.length
            ? attack.damageRows.join(" / ")
            : `${detail} ${attack.special || ""}`;
          const damageParts = this._extractDamagePartsFromText(damageSource);
          const damageFormula = damageParts[0]?.formula || this._parseRawDamageFormula(attack.damage);
          const damageType = damageParts.map(part => part.type).filter(Boolean).join(" / ");
          const accuracyBreakdown = this._formatSignedNumber(accuracyValue);
          const comment = collectNotesForBlock(attack);
          return {
            id: randomID(),
            type: "attacks",
            sheet: coreSheet.properties.title,
            keyword: attack.attackName,
            attackName: attack.attackName,
            attackDetail: detail,
            damageType,
            damageParts,
            value: accuracyValue,
            accuracyBase: accuracyValue,
            accuracyAbility: "",
            accuracyAbilityMod: 0,
            accuracyProficiency: 0,
            accuracyBreakdown,
            critRange: Number.isFinite(critRange) ? critRange : 20,
            damage: damageFormula || "",
            comment,
            attackNameCell: "",
            critRangeCell: "",
            damageCell: ""
          };
        })
        .filter(attack => attack.damage);

      if (!attacks.length) throw new Error(`No attacks detected in sheet "${coreSheet.properties.title}".`);

      const refs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
      const existingAttackCount = refs.filter(r => r?.type === "attacks").length;
      await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, refs);

      const nonAttacks = refs.filter(r => r?.type !== "attacks");
      await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, [...nonAttacks, ...attacks]);
      ui.notifications.info(`Updated attacks from ${coreSheet.properties.title}: ${existingAttackCount} → ${attacks.length}. Use Undo Bulk Add to revert.`);
      this.render(false);
    } catch (error) {
      ui.notifications.error(`Update attacks failed: ${error.message}`);
    }
  }

  async _onBulkAddSaves(event) {
    event.preventDefault();

    const sheetNames = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_NAMES) || [];
    const defaultSheet = sheetNames[0] || "Core";
    const options = sheetNames.map(name => `<option value="${name}">${name}</option>`).join("");

    const content = `
      <form class="sheexcel-bulk-saves-form" style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
        <label>Sheet</label>
        <select name="sheet">${options || `<option value="${defaultSheet}">${defaultSheet}</option>`}</select>

        <label>Top Left Cell</label>
        <input name="topLeftCell" type="text" value="A7" />

        <label>Bottom Right Cell</label>
        <input name="bottomRightCell" type="text" value="D12" />
      </form>
      <div style="margin-top:10px;padding:8px 10px;border:1px solid rgba(191,160,90,0.35);border-radius:6px;background:rgba(191,160,90,0.08);">
        <p style="margin:0 0 6px 0;font-weight:600;">Saves scanner</p>
        <p style="margin:2px 0;">Scans all <code>*Save:</code> + value pairs in this rectangle and creates saves references.</p>
      </div>
    `;

    new Dialog({
      title: "Bulk Add Saves",
      content,
      buttons: {
        add: {
          label: "Add",
          callback: async (html) => {
            try {
              const form = html[0].querySelector(".sheexcel-bulk-saves-form");
              const data = new FormData(form);
              const sheet = String(data.get("sheet") || defaultSheet);
              const topLeftCell = String(data.get("topLeftCell") || "A7");
              const bottomRightCell = String(data.get("bottomRightCell") || "D12");

              const scan = await this._scanArea(sheet, topLeftCell, bottomRightCell);
              const saveLabels = new Set([
                "strength", "dexterity", "constitution", "intelligence", "wisdom", "charisma",
                "str", "dex", "con", "int", "wis", "cha",
                "fortitude", "reflex", "will"
              ]);
              const normalizeSaveLabel = (value) => String(value || "")
                .toLowerCase()
                .replace(/[^a-z\s]/g, " ")
                .replace(/\s+/g, " ")
                .trim();
              const looksLikeSave = ({ normalized, label }) => {
                const clean = normalizeSaveLabel(normalized || label);
                if (!clean) return false;
                if (clean.includes("save")) return true;
                if (saveLabels.has(clean)) return true;
                return clean.split(" ").some(part => saveLabels.has(part));
              };

              const entries = this._extractLabelValueEntries(
                scan,
                (entry) => looksLikeSave(entry),
                { requireNumericValue: true }
              );
              if (!entries.length) throw new Error("No saves detected in selected range.");

              const existingRefs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
              await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, existingRefs);

              const seen = new Set();
              const additions = [];
              for (const entry of entries) {
                const cell = this._buildSingleCell(entry.valueColumn, entry.row);
                const key = `${entry.label}|${cell}|${sheet}`;
                if (seen.has(key)) continue;
                seen.add(key);

                additions.push({
                  id: randomID(),
                  cell,
                  sheet,
                  keyword: entry.label,
                  type: "saves",
                  value: "",
                  attackNameCell: "",
                  critRangeCell: "",
                  damageCell: "",
                  subchecks: []
                });
              }

              if (!additions.length) throw new Error("No saves generated from detected entries.");

              await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, [...existingRefs, ...additions]);
              ui.notifications.info(`Added ${additions.length} saves. Use Undo Bulk Add to revert.`);
              this.render(false);
            } catch (error) {
              ui.notifications.error(`Bulk saves failed: ${error.message}`);
            }
          }
        },
        cancel: { label: "Cancel" }
      },
      default: "add"
    }).render(true);
  }

  _extractGearCurrency(scan) {
    const currency = {
      onPerson: { gold: "", silver: "", copper: "" },
      banked: { gold: "", silver: "", copper: "" }
    };
    const normalize = (value) => String(value || "").trim().toLowerCase();
    const labelMap = {
      "gold:": "gold",
      "silver:": "silver",
      "copper:": "copper"
    };

    const maxRow = Math.min(scan.endRow, scan.startRow + 30);
    let bankedStartCol = null;

    for (let row = scan.startRow; row <= maxRow; row++) {
      // Detect the Banked column marker
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const cell = normalize(scan.readCell(row, col));
        if (cell === "banked") {
          bankedStartCol = col;
          break;
        }
      }

      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const cell = normalize(scan.readCell(row, col));
        const currencyKey = labelMap[cell];
        if (!currencyKey) continue;

        // Get next non-empty cell to the right as value
        let value = "";
        for (let vc = col + 1; vc <= scan.endColumn; vc++) {
          const raw = String(scan.readCell(row, vc) || "").trim();
          if (!raw) continue;
          value = raw;
          break;
        }
        if (!value) continue;

        const target = bankedStartCol !== null && col >= bankedStartCol
          ? currency.banked
          : currency.onPerson;
        target[currencyKey] = value;
      }
    }

    const hasValues = Object.values(currency.onPerson).some(Boolean)
      || Object.values(currency.banked).some(Boolean);
    return hasValues ? currency : null;
  }

  _extractGearBlocks(scan) {
    const normalizeLabel = (value) => String(value || "")
      .trim()
      .replace(/:$/, "")
      .toLowerCase();
    const labelMap = {
      "type": "gearType",
      "description": "description",
      "value": "value",
      "rarity": "rarity",
      "legality": "legality",
      "power": "power",
      "bulk": "bulk",
      "noise": "noise",
      "weight": "weight",
      "durability": "durability",
      "integrity": "integrity",
      "resilience": "resilience",
      "reach": "reach",
      "draw": "draw",
      "accuracy": "accuracy",
      "critical": "critical",
      "damage": "damage",
      "armor": "armor",
      "ability": "abilities"
    };
    const multiFields = new Set(["description", "abilities"]);

    const gears = [];
    const numberPattern = /^\d+$/;
    const rowHasAnyValue = (row) => {
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        if (scan.readCell(row, col)) return true;
      }
      return false;
    };
    const isGearNameRow = (row) => {
      const firstCol = String(scan.readCell(row, scan.startColumn) || "").trim();
      const secondCol = String(scan.readCell(row, scan.startColumn + 1) || "").trim();
      if (!numberPattern.test(firstCol)) return false;
      if (!secondCol || !/[A-Za-z]/.test(secondCol)) return false;
      return !normalizeLabel(secondCol).includes("type");
    };

    for (let row = scan.startRow; row <= scan.endRow; row++) {
      if (!isGearNameRow(row)) continue;

      const firstCol = String(scan.readCell(row, scan.startColumn)).trim();
      const gearName = String(scan.readCell(row, scan.startColumn + 1)).trim();

      const blockEnd = (() => {
        for (let r = row + 1; r <= scan.endRow; r++) {
          if (isGearNameRow(r)) return r - 1;
          if (!rowHasAnyValue(r) && r > row + 100) return r - 1;
        }
        return scan.endRow;
      })();

      const gear = {
        id: randomID(),
        type: "gears",
        sheet: "",
        keyword: gearName,
        gearName,
        quantity: firstCol,
        gearNameCell: this._buildSingleCell(scan.startColumn + 1, row),
        gearType: "",
        description: "",
        value: "",
        rarity: "",
        legality: "",
        power: "",
        bulk: "",
        noise: "",
        weight: "",
        durability: "",
        integrity: "",
        resilience: "",
        reach: "",
        draw: "",
        accuracy: "",
        critical: "",
        damage: "",
        armor: "",
        abilities: "",
        subchecks: []
      };

      const multiData = {};
      let activeMulti = null;

      for (let r = row + 1; r <= blockEnd; r++) {
        const rowData = [];
        for (let col = scan.startColumn; col <= scan.endColumn; col++) {
          rowData.push(scan.readCell(r, col));
        }

        // Check for label: value pattern in this row
        let foundLabel = false;
        // Search all columns for labels (ending with :)
        for (let col = scan.startColumn; col <= scan.endColumn; col++) {
          const cell = String(rowData[col - scan.startColumn] || "").trim();
          if (!cell.endsWith(":")) continue;

          foundLabel = true;
          activeMulti = null;

          const normalized = normalizeLabel(cell);
          const field = labelMap[normalized];
          if (!field) continue;

          // Collect all non-empty cells after the label in this row
          const values = [];
          let firstValueCol = null;
          for (let vc = col + 1; vc <= scan.endColumn; vc++) {
            const val = String(rowData[vc - scan.startColumn] || "").trim();
            if (!val) continue;
            if (!firstValueCol) firstValueCol = vc;
            values.push(val);
          }

          if (!firstValueCol || !values.length) {
            // Even if no values in this row, start multi-field capture for abilities
            if (field === "abilities") {
              activeMulti = {
                field: field,
                col: col + 1,
                endCol: scan.endColumn,
                startRow: r,
                endRow: r,
                values: []
              };
              if (!multiData[field]) multiData[field] = activeMulti;
            }
            continue;
          }

          const value = values.join(" ");
          if (multiFields.has(field)) {
            activeMulti = {
              field: field,
              col: firstValueCol,
              endCol: scan.endColumn,
              startRow: r,
              endRow: r,
              values: value ? [value] : []
            };
            if (!multiData[field]) multiData[field] = activeMulti;
          } else {
            gear[field] = value;
            const range = firstValueCol === scan.endColumn
              ? this._buildSingleCell(firstValueCol, r)
              : this._buildAbsoluteRange(firstValueCol, scan.endColumn, r);
            gear[`${field}Cell`] = range;
          }
          break; // Process first label in row
        }

        // If this row continues a multi-field, append to it
        if (!foundLabel && activeMulti) {
          const continuationValues = [];
          for (let col = scan.startColumn; col <= scan.endColumn; col++) {
            const val = String(scan.readCell(r, col) || "").trim();
            if (val) continuationValues.push(val);
          }
          if (continuationValues.length) {
            activeMulti.endRow = r;
            // For abilities, preserve as separate items; for others join with space
            const joinValue = activeMulti.field === "abilities" 
              ? continuationValues.join("\n") 
              : continuationValues.join(" ");
            activeMulti.values.push(joinValue);
          } else if (rowHasAnyValue(r)) {
            // Continue multi if row has any value
          } else {
            activeMulti = null;
          }
        }
      }

      // Finalize multi-field values
      for (const [field, data] of Object.entries(multiData)) {
        if (!data.values.length) continue;
        const joiner = "\n";
        let text = data.values.join(joiner);
        
        // Clean up abilities: remove all tabs, non-breaking spaces, and excessive whitespace
        if (field === "abilities") {
          // Replace tabs and non-breaking spaces with regular spaces
          text = text.replace(/[\t\u00A0]/g, " ");
          // Trim each line and collapse multiple spaces
          text = text.split("\n")
            .map(line => line.trim().replace(/  +/g, " "))
            .filter(line => line.length > 0)  // Remove empty lines
            .join("\n");
          // Format ability labels (text before :) as bold
          text = text.replace(/^([^:\n]*?):(\s|$)/gm, "<strong>$1:</strong>$2");
        }
        
        gear[field] = text;
        const range = (data.startRow === data.endRow && data.col === scan.endColumn)
          ? this._buildSingleCell(data.col, data.startRow)
          : this._buildRectRange(data.col, data.startRow, scan.endColumn, data.endRow);
        gear[`${field}Cell`] = range;
      }

      // No need to extract gearPrimaryType - prepareData will categorize by keywords
      gears.push(gear);
    }

    return gears;
  }

  async _onUpdateGearsFromSheet(event) {
    event.preventDefault();

    try {
      this._invalidateApiCache();
      const metadata = await this._fetchSheetMetadata();
      const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
      if (!sheets.length) throw new Error("No sheet metadata found.");

      const gearSheet = sheets.find(s => /gear|equipment/i.test(s.properties?.title || "")) || sheets[0];
      const sheetName = gearSheet?.properties?.title;
      if (!sheetName) throw new Error("Could not determine sheet name for gears.");

      const rowCount = Math.max(2, Number(gearSheet.properties?.gridProperties?.rowCount) || 500);
      const columnCount = Math.max(2, Number(gearSheet.properties?.gridProperties?.columnCount) || 11);
      const bottomRight = `${this._columnNumberToLetters(columnCount)}${rowCount}`;

      const scan = await this._scanArea(sheetName, "A1", bottomRight);
      const gearCurrency = this._extractGearCurrency(scan);
      const gears = this._extractGearBlocks(scan);
      if (!gears.length) throw new Error(`No gear detected in sheet \"${sheetName}\".`);

      const refs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
      const existingGearCount = refs.filter(r => r?.type === "gears").length;
      await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, refs);

      const nonGears = refs.filter(r => r?.type !== "gears");
      const additions = gears.map(gear => ({ ...gear, sheet: sheetName }));

      await this.actor.setFlag(MODULE_NAME, FLAGS.GEAR_CURRENCY, gearCurrency);

      await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, [...nonGears, ...additions]);
      ui.notifications.info(`Updated gears from ${sheetName}: ${existingGearCount} → ${additions.length}. Use Undo Bulk Add to revert.`);
      this.render(false);
    } catch (error) {
      ui.notifications.error(`Update gears failed: ${error.message}`);
    }
  }

  async _onBulkAddSpells(event) {
    event.preventDefault();

    const sheetNames = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_NAMES) || [];
    const defaultSheet = sheetNames[0] || "Core";
    const options = sheetNames.map(name => `<option value="${name}">${name}</option>`).join("");

    const content = `
      <form class="sheexcel-bulk-spells-form" style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
        <label>Sheet</label>
        <select name="sheet">${options || `<option value="${defaultSheet}">${defaultSheet}</option>`}</select>

        <label>Top Left Cell</label>
        <input name="topLeftCell" type="text" value="A1" />

        <label>Bottom Right Cell</label>
        <input name="bottomRightCell" type="text" value="K200" />
      </form>
      <div style="margin-top:10px;padding:8px 10px;border:1px solid rgba(191,160,90,0.35);border-radius:6px;background:rgba(191,160,90,0.08);">
        <p style="margin:0 0 6px 0;font-weight:600;">Spells scanner</p>
        <p style="margin:2px 0;">Scans spell blocks where the name is followed by labeled rows like <code>Circle:</code>, <code>Type:</code>, <code>Components:</code>.</p>
        <p style="margin:2px 0;">Description, Effect, Empower, and Components can span multiple rows.</p>
      </div>
    `;

    new Dialog({
      title: "Bulk Add Spells",
      content,
      buttons: {
        add: {
          label: "Add",
          callback: async (html) => {
            try {
              const form = html[0].querySelector(".sheexcel-bulk-spells-form");
              const data = new FormData(form);
              const sheet = String(data.get("sheet") || defaultSheet);
              const topLeftCell = String(data.get("topLeftCell") || "A1");
              const bottomRightCell = String(data.get("bottomRightCell") || "K200");

              const scan = await this._scanArea(sheet, topLeftCell, bottomRightCell);
              const spells = this._extractSpellBlocks(scan);
              if (!spells.length) throw new Error("No spells detected in selected range.");

              const existingRefs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
              await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, existingRefs);

              const additions = spells.map(spell => ({
                ...spell,
                sheet
              }));

              await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, [...existingRefs, ...additions]);
              ui.notifications.info(`Added ${additions.length} spells. Use Undo Bulk Add to revert.`);
              this.render(false);
            } catch (error) {
              ui.notifications.error(`Bulk spells failed: ${error.message}`);
            }
          }
        },
        cancel: { label: "Cancel" }
      },
      default: "add"
    }).render(true);
  }

  async _onUpdateSpellsFromSheet(event) {
    event.preventDefault();

    try {
      this._invalidateApiCache();
      const metadata = await this._fetchSheetMetadata();
      const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
      if (!sheets.length) throw new Error("No sheet metadata found.");

      const spellSheet = sheets.find(s => /spells/i.test(s.properties?.title || "")) || sheets[0];
      const sheetName = spellSheet?.properties?.title;
      if (!sheetName) throw new Error("Could not determine sheet name for spells.");

      const rowCount = Math.max(2, Number(spellSheet.properties?.gridProperties?.rowCount) || 200);
      const columnCount = Math.max(2, Number(spellSheet.properties?.gridProperties?.columnCount) || 11);
      const bottomRight = `${this._columnNumberToLetters(columnCount)}${rowCount}`;

      const scan = await this._scanArea(sheetName, "A1", bottomRight);
      const spells = this._extractSpellBlocks(scan);
      if (!spells.length) throw new Error(`No spells detected in sheet \"${sheetName}\".`);

      const refs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
      const existingSpellCount = refs.filter(r => r?.type === "spells").length;
      await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, refs);

      const nonSpells = refs.filter(r => r?.type !== "spells");
      const additions = spells.map(spell => ({ ...spell, sheet: sheetName }));

      await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, [...nonSpells, ...additions]);
      ui.notifications.info(`Updated spells from ${sheetName}: ${existingSpellCount} → ${additions.length}. Use Undo Bulk Add to revert.`);
      this.render(false);
    } catch (error) {
      ui.notifications.error(`Update spells failed: ${error.message}`);
    }
  }

  async _onUpdateAbilitiesFromSheet(event) {
    event.preventDefault();

    try {
      console.log("Sheexcel | Update abilities start");
      this._invalidateApiCache();
      const metadata = await this._fetchSheetMetadata();
      const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
      if (!sheets.length) throw new Error("No sheet metadata found.");

      console.log("Sheexcel | Ability update metadata sheets", sheets.map(s => s?.properties?.title || ""));

      const abilitySheet = sheets.find(s => /^abilities?$/i.test(s.properties?.title || ""))
        || sheets.find(s => /abilities?|ability/i.test(s.properties?.title || ""))
        || sheets[0];
      const sheetName = abilitySheet?.properties?.title;
      if (!sheetName) throw new Error("Could not determine sheet name for abilities.");

      const rowCount = Math.max(2, Number(abilitySheet.properties?.gridProperties?.rowCount) || 300);
      const columnCount = Math.max(2, Number(abilitySheet.properties?.gridProperties?.columnCount) || 12);
      const bottomRight = `${this._columnNumberToLetters(columnCount)}${rowCount}`;

      console.log("Sheexcel | Ability update target", {
        sheetName,
        rowCount,
        columnCount,
        range: `A1:${bottomRight}`
      });

      const scan = await this._scanArea(sheetName, "A1", bottomRight);
      const previewRows = [];
      for (let row = scan.startRow; row <= Math.min(scan.endRow, scan.startRow + 7); row++) {
        const cells = [];
        for (let col = scan.startColumn; col <= Math.min(scan.endColumn, scan.startColumn + 5); col++) {
          const v = scan.readCell(row, col);
          if (v) cells.push(v);
        }
        previewRows.push({ row, cells });
      }
      console.log("Sheexcel | Ability scan preview", previewRows);

      const profile = {
        totalRows: scan.endRow - scan.startRow + 1,
        rowsWithValues: 0,
        colUsage: {},
        firstNonEmptyColUsage: {},
        headerCodeTitleRows: []
      };
      for (let row = scan.startRow; row <= scan.endRow; row++) {
        const rowCells = [];
        for (let col = scan.startColumn; col <= scan.endColumn; col++) {
          const value = scan.readCell(row, col);
          if (!value) continue;
          rowCells.push({ col, value: String(value).trim() });
          profile.colUsage[col] = (profile.colUsage[col] || 0) + 1;
        }
        if (!rowCells.length) continue;

        profile.rowsWithValues += 1;
        const firstCol = rowCells[0].col;
        profile.firstNonEmptyColUsage[firstCol] = (profile.firstNonEmptyColUsage[firstCol] || 0) + 1;

        const first = rowCells[0];
        const second = rowCells[1];
        if (/^[A-Z]{2,4}$/.test(first.value) && second?.value && /[A-Za-z]/.test(second.value)) {
          if (profile.headerCodeTitleRows.length < 30) {
            profile.headerCodeTitleRows.push({ row, code: first.value, title: second.value });
          }
        }
      }
      console.log("Sheexcel | Ability scan profile", profile);

      const abilities = this._extractAbilityBlocks(scan);
      if (!abilities.length) throw new Error(`No abilities detected in sheet \"${sheetName}\".`);

      console.log("Sheexcel | Ability update parsed count", {
        count: abilities.length,
        names: abilities.slice(0, 10).map(a => a.abilityName || a.keyword)
      });

      const refs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
      const existingAbilityCount = refs.filter(r => r?.type === "abilities").length;
      await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, refs);

      const nonAbilities = refs.filter(r => r?.type !== "abilities");
      const additions = abilities.map(ability => ({ ...ability, sheet: sheetName }));

      await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, [...nonAbilities, ...additions]);
      ui.notifications.info(`Updated abilities from ${sheetName}: ${existingAbilityCount} → ${additions.length}. Use Undo Bulk Add to revert.`);
      console.log("Sheexcel | Update abilities success", {
        sheetName,
        previous: existingAbilityCount,
        next: additions.length
      });
      this.render(false);
    } catch (error) {
      console.error("Sheexcel | Update abilities failed", error);
      ui.notifications.error(`Update abilities failed: ${error.message}`);
    }
  }

  _extractRestEntries(scan) {
    const rows = [];

    for (let row = scan.startRow; row <= scan.endRow; row++) {
      const cells = [];
      for (let col = scan.startColumn; col <= scan.endColumn; col++) {
        const text = String(scan.readCell(row, col) || "").trim();
        if (!text) continue;
        cells.push({ col, text });
      }

      if (!cells.length) continue;
      rows.push({ row, cells, firstCol: cells[0].col, text: cells.map((cell) => cell.text).join(" ") });
    }

    if (!rows.length) return [];

    const indentationCols = Array.from(new Set(rows.map((row) => row.firstCol))).sort((a, b) => a - b);
    const indentationRank = new Map(indentationCols.map((col, index) => [col, index]));
    const baseCol = indentationCols[0];
    const entries = [];
    let currentEntry = null;

    rows.forEach((rowData) => {
      const level = indentationRank.get(rowData.firstCol) ?? 0;

      if (level === 0 || rowData.firstCol === baseCol) {
        currentEntry = {
          id: randomID(),
          section: "Rest",
          title: rowData.text,
          summary: "",
          details: [],
          row: rowData.row
        };
        entries.push(currentEntry);
        return;
      }

      if (!currentEntry) {
        currentEntry = {
          id: randomID(),
          section: "Rest",
          title: rowData.text,
          summary: "",
          details: [],
          row: rowData.row
        };
        entries.push(currentEntry);
        return;
      }

      currentEntry.details.push({
        text: rowData.text,
        level: Math.max(1, level)
      });
    });

    return entries;
  }

  async _onUpdateRestFromSheet(event) {
    event.preventDefault();

    try {
      this._invalidateApiCache();
      const metadata = await this._fetchSheetMetadata();
      const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
      if (!sheets.length) throw new Error("No sheet metadata found.");

      const restSheet = sheets.find((sheet) => /^rest$/i.test(sheet.properties?.title || ""))
        || sheets.find((sheet) => /rest/i.test(sheet.properties?.title || ""));
      const sheetName = restSheet?.properties?.title;
      if (!sheetName) throw new Error("Could not find a sheet named Rest.");

      const rowCount = Math.max(2, Number(restSheet.properties?.gridProperties?.rowCount) || 300);
      const columnCount = Math.max(2, Number(restSheet.properties?.gridProperties?.columnCount) || 12);
      const bottomRight = `${this._columnNumberToLetters(columnCount)}${rowCount}`;

      const scan = await this._scanArea(sheetName, "A1", bottomRight);
      const entries = this._extractRestEntries(scan);
      if (!entries.length) throw new Error(`No rest entries detected in sheet \"${sheetName}\".`);

      await this.actor.setFlag(MODULE_NAME, FLAGS.REST_ENTRIES, entries);
      ui.notifications.info(`Updated rest entries from ${sheetName}: ${entries.length} loaded.`);
      this.render(false);
    } catch (error) {
      ui.notifications.error(`Update rest failed: ${error.message}`);
    }
  }

  async _onPostRestLineToChat(entryIndex, detailIndex = null, lineType = "detail") {
    const entries = this.actor.getFlag(MODULE_NAME, FLAGS.REST_ENTRIES) || [];
    const entry = entries[entryIndex];
    if (!entry) return;

    const title = String(entry.title || "Rest").trim() || "Rest";
    const escapedTitle = escapeHtml(title);

    if (lineType === "card") {
      const summary = String(entry.summary || "").trim();
      const details = Array.isArray(entry.details) ? entry.details : [];
      const summaryHtml = summary
        ? `<div class="sheexcel-rest-summary">${enrichTextWithInlineRolls(summary, { actorId: this.actor?.id || "", contextLabel: title })}</div>`
        : "";
      const detailHtml = details
        .map((detail) => {
          const detailText = typeof detail === "string" ? detail.trim() : String(detail?.text || "").trim();
          if (!detailText) return "";
          const level = typeof detail === "string" ? 1 : Math.max(1, Number(detail?.level) || 1);
          return `<div class="sheexcel-rest-detail-line sheexcel-rest-detail-level-${level}">${enrichTextWithInlineRolls(detailText, { actorId: this.actor?.id || "", contextLabel: title })}</div>`;
        })
        .filter(Boolean)
        .join("");

      if (!summaryHtml && !detailHtml) return;

      const content = `
        <div class="sheexcel-spell-chat sheexcel-rest-chat">
          <div class="sheexcel-spell-chat-sub">${escapedTitle}</div>
          <div class="sheexcel-spell-chat-sections">
            ${summaryHtml}
            ${detailHtml ? `<div class="sheexcel-rest-details">${detailHtml}</div>` : ""}
          </div>
        </div>
      `;

      await ChatMessage.create({
        speaker: ChatMessage.getSpeaker({ actor: this.actor }),
        content
      });
      return;
    }

    let lineText = "";
    if (lineType === "summary") {
      lineText = String(entry.summary || "").trim();
    } else {
      const detail = Array.isArray(entry.details) ? entry.details[detailIndex] : null;
      if (typeof detail === "string") {
        lineText = detail.trim();
      } else {
        lineText = String(detail?.text || "").trim();
      }
    }

    if (!lineText) return;

    const content = `
      <div class="sheexcel-spell-chat sheexcel-rest-chat">
        <div class="sheexcel-spell-chat-sub">${escapedTitle}</div>
        <div class="sheexcel-spell-chat-sections">${enrichTextWithInlineRolls(lineText, { actorId: this.actor?.id || "", contextLabel: title })}</div>
      </div>
    `;

    await ChatMessage.create({
      speaker: ChatMessage.getSpeaker({ actor: this.actor }),
      content
    });
  }

  async _onUpdateAllFromSheet(event) {
    event.preventDefault();
    const noopEvent = { preventDefault() {} };

    await this._onUpdateSkillsFromSheet(noopEvent);
    await this._onUpdateAbilitiesFromSheet(noopEvent);
    await this._onUpdateSavesFromSheet(noopEvent);
    await this._onUpdateAttacksFromSheet(noopEvent);
    await this._onUpdateSpellsFromSheet(noopEvent);
    await this._onUpdateGearsFromSheet(noopEvent);
    await this._onUpdateRestFromSheet(noopEvent);

    ui.notifications.info("Update all complete.");
  }

  async _fetchSheetMetadata() {
    const cacheTtl = 30000;
    if (this._metadataCache && (Date.now() - this._metadataCache.fetchedAt) < cacheTtl) {
      return this._metadataCache.data;
    }

    const sheetId = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_ID);
    const apiKey = game.settings.get(MODULE_NAME, SETTINGS.GOOGLE_API_KEY);

    if (!sheetId) throw new Error("Missing sheet ID. Update Sheet first.");
    if (!apiKey) throw new Error("Missing Google API key.");

    const fields = "sheets(properties(title,gridProperties(rowCount,columnCount)),merges)";
    const url = `${API_CONFIG.BASE_URL}/${sheetId}?fields=${encodeURIComponent(fields)}&key=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      if (response.status === 429 && this._metadataCache?.data) {
        return this._metadataCache.data;
      }
      throw new Error(`Metadata fetch failed (${response.status})`);
    }

    const data = await response.json();
    this._metadataCache = { data, fetchedAt: Date.now() };
    return data;
  }

  async _onPostSpellToChat(index) {
    const refs = this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || [];
    const ref = refs[index];
    if (!ref || ref.type !== "spells") return;

    const toLines = (value) => String(value || "")
      .trim();

    const name = ref.spellName || ref.keyword || "Spell";
    const header = ref.circle ? `Circle ${ref.circle}` : "";

    const tags = [ref.spellType, ref.source, ref.discipline, ref.components]
      .filter(Boolean)
      .map(t => `<span class="sheexcel-spell-chat-tag">${t}</span>`)
      .join(" ");

    const stats = [
      ref.castTime ? `<div><strong>Cast:</strong> ${ref.castTime}</div>` : "",
      ref.cost ? `<div><strong>Cost:</strong> ${ref.cost}</div>` : "",
      ref.range ? `<div><strong>Range:</strong> ${ref.range}</div>` : "",
      ref.duration ? `<div><strong>Duration:</strong> ${ref.duration}</div>` : ""
    ].filter(Boolean).join("");

    const sections = [
      ref.description ? `<div><strong>Description:</strong><br>${enrichTextWithInlineRolls(toLines(ref.description), { actorId: this.actor?.id || "", contextLabel: name })}</div>` : "",
      ref.effect ? `<div><strong>Effect:</strong><br>${enrichTextWithInlineRolls(toLines(ref.effect), { actorId: this.actor?.id || "", contextLabel: name })}</div>` : "",
      ref.empower ? `<div><strong>Empower:</strong><br>${enrichTextWithInlineRolls(toLines(ref.empower), { actorId: this.actor?.id || "", contextLabel: name })}</div>` : ""
    ].filter(Boolean).join("<br>");

    const content = `
      <div class="sheexcel-spell-chat">
        <div class="sheexcel-spell-chat-title">${name}</div>
        ${header ? `<div class="sheexcel-spell-chat-sub">${header}</div>` : ""}
        ${tags ? `<div class="sheexcel-spell-chat-tags">${tags}</div>` : ""}
        ${stats ? `<div class="sheexcel-spell-chat-stats">${stats}</div>` : ""}
        ${sections ? `<div class="sheexcel-spell-chat-sections">${sections}</div>` : ""}
      </div>
    `;

    await ChatMessage.create({
      speaker: ChatMessage.getSpeaker({ actor: this.actor }),
      content
    });
  }

  async _onPostAbilityToChat(index) {
    const refs = this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || [];
    const ref = refs[index];
    if (!ref || ref.type !== "abilities") return;

    const toLines = (value) => String(value || "")
      .trim();

    const name = ref.abilityName || ref.keyword || "Ability";
    const tags = [
      ref.category ? `Category: ${ref.category}` : "",
      ref.cost ? `Cost: ${ref.cost}` : "",
      ref.target ? `Target: ${ref.target}` : "",
      ref.range ? `Range: ${ref.range}` : "",
      ref.duration ? `Duration: ${ref.duration}` : "",
      ref.trigger ? `Trigger: ${ref.trigger}` : ""
    ].filter(Boolean)
      .map(t => `<span class="sheexcel-spell-chat-tag">${t}</span>`)
      .join(" ");

    const sections = [
      ref.description ? `<div><strong>Description:</strong><br>${enrichTextWithInlineRolls(toLines(ref.description), { actorId: this.actor?.id || "", contextLabel: name })}</div>` : "",
      ref.effect ? `<div><strong>Effect:</strong><br>${enrichTextWithInlineRolls(toLines(ref.effect), { actorId: this.actor?.id || "", contextLabel: name })}</div>` : "",
      ref.notes ? `<div><strong>Notes:</strong><br>${enrichTextWithInlineRolls(toLines(ref.notes), { actorId: this.actor?.id || "", contextLabel: name })}</div>` : ""
    ].filter(Boolean).join("<br>");

    const content = `
      <div class="sheexcel-spell-chat sheexcel-ability-chat">
        <div class="sheexcel-spell-chat-title">${name}</div>
        ${tags ? `<div class="sheexcel-spell-chat-tags">${tags}</div>` : ""}
        ${sections ? `<div class="sheexcel-spell-chat-sections">${sections}</div>` : ""}
      </div>
    `;

    await ChatMessage.create({
      speaker: ChatMessage.getSpeaker({ actor: this.actor }),
      content
    });
  }

  async _getMergedAttackBlocks(sheetName, nameRow, startColumn, maxBlocks) {
    const metadata = await this._fetchSheetMetadata();
    const targetSheet = metadata.sheets?.find(s => s.properties?.title === sheetName);
    if (!targetSheet) throw new Error(`Sheet \"${sheetName}\" not found in metadata`);

    const rowIndex = nameRow - 1;
    const merges = Array.isArray(targetSheet.merges) ? targetSheet.merges : [];

    const blocks = merges
      .filter(m => m.startRowIndex <= rowIndex && m.endRowIndex > rowIndex)
      .map(m => ({
        startColumn: m.startColumnIndex + 1,
        width: m.endColumnIndex - m.startColumnIndex
      }))
      .filter(m => m.width > 1 && m.startColumn >= startColumn)
      .sort((a, b) => a.startColumn - b.startColumn);

    const deduped = [];
    const seen = new Set();
    for (const block of blocks) {
      if (seen.has(block.startColumn)) continue;
      seen.add(block.startColumn);
      deduped.push(block);
    }

    if (!deduped.length) {
      throw new Error("No merged blocks found on the Name Row at/after Start Column");
    }

    return deduped.slice(0, maxBlocks);
  }

  _onBulkAddAttacks(event) {
    event.preventDefault();

    const sheetNames = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_NAMES) || [];
    const defaultSheet = sheetNames[0] || "Core";
    const options = sheetNames.map(name => `<option value="${name}">${name}</option>`).join("");

    const content = `
      <form class="sheexcel-bulk-attack-form" style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
        <label>Sheet</label>
        <select name="sheet">${options || `<option value="${defaultSheet}">${defaultSheet}</option>`}</select>

        <label>Top Left Cell</label>
        <input name="topLeftCell" type="text" value="S4" />

        <label>Bottom Right Cell</label>
        <input name="bottomRightCell" type="text" value="AL16" />
      </form>
      <div style="margin-top:10px;padding:8px 10px;border:1px solid rgba(191,160,90,0.35);border-radius:6px;background:rgba(191,160,90,0.08);">
        <p style="margin:0 0 6px 0;font-weight:600;">Field guide</p>
        <p style="margin:2px 0;"><strong>Top Left / Bottom Right</strong>: define the full rectangle that contains your attack blocks.</p>
        <p style="margin:2px 0;"><strong>Scanner behavior</strong>: it looks for <code>Accuracy:</code> rows and auto-builds matching Value/Critical/Damage ranges inside that rectangle.</p>
        <p style="margin:2px 0;"><strong>Name source</strong>: uses the row two lines above each Accuracy row as attack name row.</p>
        <p style="margin:2px 0;"><strong>Safe to test</strong>: this stores a backup first, and <strong>Undo Bulk Add</strong> restores previous references.</p>
      </div>
      <p style="margin-top:8px;opacity:0.9;">This creates attack references and stores a backup so you can undo.</p>
    `;

    new Dialog({
      title: "Bulk Add Attack References",
      content,
      buttons: {
        add: {
          label: "Add",
          callback: async (html) => {
            try {
              const form = html[0].querySelector(".sheexcel-bulk-attack-form");
              const data = new FormData(form);

              const sheet = String(data.get("sheet") || defaultSheet);
              const topLeft = this._parseA1Cell(data.get("topLeftCell"));
              const bottomRight = this._parseA1Cell(data.get("bottomRightCell"));

              const startColumn = Math.min(topLeft.column, bottomRight.column);
              const endColumn = Math.max(topLeft.column, bottomRight.column);
              const startRow = Math.min(topLeft.row, bottomRight.row);
              const endRow = Math.max(topLeft.row, bottomRight.row);

              if (startRow === endRow || startColumn === endColumn) {
                throw new Error("Range must contain multiple rows and columns");
              }

              const safeSheet = sheet.match(/[^A-Za-z0-9_]/)
                ? `'${sheet.replace(/'/g, "''")}'`
                : sheet;

              const areaRange = `${safeSheet}!${this._columnNumberToLetters(startColumn)}${startRow}:${this._columnNumberToLetters(endColumn)}${endRow}`;
              const sheetId = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_ID);
              if (!sheetId) throw new Error("Missing sheet ID. Update Sheet first.");

              const scanJson = await apiCache.batchGet(sheetId, [areaRange]);
              const matrix = scanJson.valueRanges?.[0]?.values || [];

              const readCell = (absoluteRow, absoluteColumn) => {
                const rowOffset = absoluteRow - startRow;
                const colOffset = absoluteColumn - startColumn;
                return (matrix[rowOffset]?.[colOffset] ?? "").toString().trim();
              };

              const existingRefs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
              await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, existingRefs);

              const additions = [];
              let attackCounter = 1;

              for (let row = startRow; row <= endRow; row++) {
                const accuracyCols = [];
                for (let col = startColumn; col <= endColumn; col++) {
                  const cellValue = readCell(row, col).toLowerCase();
                  if (cellValue === "accuracy:" || cellValue === "accuracy") {
                    accuracyCols.push(col);
                  }
                }

                for (let i = 0; i < accuracyCols.length; i++) {
                  const labelCol = accuracyCols[i];
                  const nextLabelCol = accuracyCols[i + 1] || (endColumn + 1);
                  const blockEndCol = Math.max(labelCol, nextLabelCol - 1);

                  let valueStartCol = null;
                  for (let col = labelCol + 1; col <= blockEndCol; col++) {
                    if (readCell(row, col)) {
                      valueStartCol = col;
                      break;
                    }
                  }
                  if (!valueStartCol) continue;

                  const nameRow = row - 2;
                  const critRow = row + 1;
                  const damageRow = row + 2;
                  if (nameRow < startRow || damageRow > endRow) continue;

                  const keyword = readCell(nameRow, labelCol) || `Attack ${attackCounter}`;

                  additions.push({
                    id: randomID(),
                    type: "attacks",
                    keyword,
                    sheet,
                    cell: this._buildAbsoluteRange(valueStartCol, blockEndCol, row),
                    attackNameCell: this._buildAbsoluteRange(labelCol, blockEndCol, nameRow),
                    critRangeCell: this._buildAbsoluteRange(valueStartCol, blockEndCol, critRow),
                    damageCell: this._buildAbsoluteRange(valueStartCol, blockEndCol, damageRow),
                    value: "",
                    subchecks: []
                  });

                  attackCounter += 1;
                }
              }

              if (!additions.length) {
                throw new Error("No attack blocks detected. Check the selected range and labels.");
              }

              await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, [...existingRefs, ...additions]);
              ui.notifications.info(`Added ${additions.length} attack references. Use Undo Bulk Add to revert.`);
              this.render(false);
            } catch (error) {
              ui.notifications.error(`Bulk add failed: ${error.message}`);
            }
          }
        },
        cancel: { label: "Cancel" }
      },
      default: "add"
    }).render(true);
  }

  async _onUndoBulkAddAttacks(event) {
    event.preventDefault();
    const backup = this.actor.getFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP);
    if (!Array.isArray(backup)) {
      ui.notifications.warn("No bulk add backup available.");
      return;
    }

    await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, backup);
    await this.actor.unsetFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP);
    ui.notifications.info("Reverted last bulk add.");
    this.render(false);
  }

  async _clearReferencesByType(type, label) {
    const refs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
    const toRemove = refs.filter(ref => ref?.type === type).length;

    if (!toRemove) {
      ui.notifications.warn(`No ${label.toLowerCase()} references found.`);
      return;
    }

    const confirmed = confirm(`Delete all ${label.toLowerCase()} references (${toRemove})? You can use Undo Bulk Add to restore.`);
    if (!confirmed) return;

    await this.actor.setFlag(MODULE_NAME, FLAGS.BULK_ADD_BACKUP, refs);
    const filtered = refs.filter(ref => ref?.type !== type);
    await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, filtered);

    ui.notifications.info(`Deleted ${toRemove} ${label.toLowerCase()} references.`);
    this.render(false);
  }

  async _onClearChecks(event) {
    event.preventDefault();
    await this._clearReferencesByType("checks", "Checks");
  }

  async _onClearSaves(event) {
    event.preventDefault();
    await this._clearReferencesByType("saves", "Saves");
  }

  async _onClearAttacks(event) {
    event.preventDefault();
    await this._clearReferencesByType("attacks", "Attacks");
  }

  async _onClearSpells(event) {
    event.preventDefault();
    await this._clearReferencesByType("spells", "Spells");
  }

  async _onClearAbilities(event) {
    event.preventDefault();
    await this._clearReferencesByType("abilities", "Abilities");
  }

  async _onClearGears(event) {
    event.preventDefault();
    await this._clearReferencesByType("gears", "Gears");
  }

  _onAddReference(event) {
    event.preventDefault();
    const container = this.element.find(".sheexcel-references-list");
    const idx = container.find(".sheexcel-reference-card").length;
    const sheets = this.actor.getFlag("sheexcel_updated", "sheetNames") || [];
    const options = sheets.map(n => `<option value="${n}">${n}</option>`).join("");
    const newId = randomID();
    
    const card = $(`
      <div class="sheexcel-reference-card" data-index="${idx}" data-id="${newId}">
        <div class="sheexcel-reference-row">
          <div class="sheexcel-input-group cell-ref">
            <input type="text" class="sheexcel-reference-input" data-type="cell" data-index="${idx}" placeholder="A1" title="Spreadsheet cell reference (e.g., A1, B5)">
          </div>
          <div class="sheexcel-input-group keyword">
            <input type="text" class="sheexcel-reference-input" data-type="keyword" data-index="${idx}" placeholder="Skill name" title="Display name for this reference">
          </div>
          <div class="sheexcel-input-group sheet">
            <select class="sheexcel-reference-input" data-type="sheet" data-index="${idx}" title="Select sheet tab">${options}</select>
          </div>
          <div class="sheexcel-input-group type">
            <select class="sheexcel-reference-input" data-type="refType" data-index="${idx}" title="Reference category">
              <option value="checks">Checks</option>
              <option value="saves">Saves</option>
              <option value="attacks">Attacks</option>
              <option value="spells">Spells</option>
            </select>
          </div>
          <div class="sheexcel-input-group value">
            <span class="sheexcel-reference-value-display" title="Current value from spreadsheet"><em>empty</em></span>
          </div>
          <div class="sheexcel-input-group actions">
            <button type="button" class="sheexcel-action-btn refresh sheexcel-reference-remove-save" data-index="${idx}" title="Refresh value from spreadsheet">
              <span aria-hidden="true">↻</span>
            </button>
            <button type="button" class="sheexcel-action-btn delete sheexcel-reference-remove-button" data-index="${idx}" title="Remove this reference">
              <span aria-hidden="true">✕</span>
            </button>
          </div>
        </div>
      </div>
    `);
    
    // If no references exist yet, replace empty state
    if (idx === 0) {
      const emptyState = this.element.find(".sheexcel-empty-state");
      if (emptyState.length) {
        emptyState.parent().html(`
          <div class="sheexcel-references-table">
            <div class="sheexcel-table-header">
              <div class="sheexcel-header-cell cell-ref">Cell</div>
              <div class="sheexcel-header-cell keyword">Name</div>
              <div class="sheexcel-header-cell sheet">Sheet</div>
              <div class="sheexcel-header-cell type">Type</div>
              <div class="sheexcel-header-cell value">Value</div>
              <div class="sheexcel-header-cell actions">Actions</div>
            </div>
            <div class="sheexcel-references-list"></div>
          </div>
        `);
        container = this.element.find(".sheexcel-references-list");
      }
    }
    
    container.append(card);
    
    // Focus on the first input for immediate editing
    card.find('input[data-type="cell"]').focus();
  }

  _onRemoveReference(event) {
    event.preventDefault();
    const $card = $(event.currentTarget).closest(".sheexcel-reference-card");
    
    // Add confirmation for destructive action
    const keyword = $card.find('input[data-type="keyword"]').val() || "this reference";
    const confirmed = confirm(`Remove ${keyword}? This cannot be undone.`);
    
    if (confirmed) {
      // Fade out the card before removing
      $card.fadeOut(300, function() {
        $(this).remove();
        
        // If no references left, show empty state
        const remainingCards = $card.parent().find(".sheexcel-reference-card").length;
        if (remainingCards === 0) {
          $card.closest(".sheexcel-references-table").replaceWith(`
            <div class="sheexcel-empty-state">
              <div class="sheexcel-empty-icon">
                <span aria-hidden="true">▦</span>
              </div>
              <h4>No References Configured</h4>
              <p>Add cell references to connect your character data with Google Sheets.</p>
            </div>
          `);
        }
      });
    }
  }

  async _onFetchAndUpdateCellValueByIndex(index) {
    const refs = foundry.utils.deepClone(this.actor.getFlag(MODULE_NAME, FLAGS.CELL_REFERENCES) || []);
    const sheetId = this.actor.getFlag(MODULE_NAME, FLAGS.SHEET_ID);
    
    if (!refs[index] || !sheetId) {
      ui.notifications.warn("Invalid reference or missing sheet ID");
      return;
    }

    const { cell, sheet } = refs[index];
    const $button = this.element.find(`[data-index="${index}"] .sheexcel-reference-remove-save`);

    try {
      // Validate inputs before making API call
      const validCell = validateCellReference(cell);
      const safeSheet = sheet.match(/[^A-Za-z0-9_]/)
        ? `'${sheet.replace(/'/g, "''")}'`
        : sheet;
      const range = `${safeSheet}!${validCell}`;

      // Use loading manager for visual feedback
      const apiPromise = (async () => {
        // Use cached API
        const json = await apiCache.batchGet(sheetId, [range]);
        return json.valueRanges?.[0]?.values?.[0]?.[0] || "";
      })();

      const value = await loadingManager.withLoading($button, `fetch-${index}`, apiPromise);
      
      refs[index].value = value;
      await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, refs);
      this.render(false);
      
      ui.notifications.info(`Updated ${refs[index].keyword}: ${value || 'empty'}`);
    } catch (error) {
      console.error("❌ Sheexcel | Cell fetch failed:", error);
      
      if (error instanceof ValidationError) {
        ui.notifications.error(`Validation error: ${error.message}`);
      } else {
        refs[index].value = "";
        await this.actor.setFlag(MODULE_NAME, FLAGS.CELL_REFERENCES, refs);
        ui.notifications.error(`Failed to fetch cell value: ${error.message}`);
      }
      
      this.render(false);
    }
  }


  _onToggleSidebar(event) {
    event.preventDefault();
    const c = !this.actor.getFlag(MODULE_NAME, FLAGS.SIDEBAR_COLLAPSED);
    this.element.find('.sheexcel-sidebar').toggleClass('collapsed', c);
    this.actor.setFlag(MODULE_NAME, FLAGS.SIDEBAR_COLLAPSED, c);
  }

  _onToggleTab(event) {
    event.preventDefault();
    const tab = event.currentTarget.dataset.tab;
    this.element.find(".sheexcel-sheet-tabs .item, .sheexcel-sidebar-tab").removeClass("active");
    this.element.find(`.sheexcel-sidebar-tab[data-tab="${tab}"]`).addClass("active");
    event.currentTarget.classList.add("active");

    const form = this.element.find("form.sheexcel-sheet")[0];
    const target = form || this.element[0];
    if (target) target.dataset.activeTab = tab;
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