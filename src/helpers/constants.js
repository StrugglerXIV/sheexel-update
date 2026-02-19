// helpers/constants.js

export const MODULE_NAME = "sheexcel_updated";

export const FLAGS = {
  SHEET_URL: "sheetUrl",
  SHEET_ID: "sheetId", 
  CELL_REFERENCES: "cellReferences",
  BULK_ADD_BACKUP: "bulkAddBackup",
  SHEET_NAMES: "sheetNames",
  GEAR_CURRENCY: "gearCurrency",
  ARMOR_CACHE: "armorCache",
  STATS_CACHE: "statsCache",
  HIDE_MENU: "hideMenu",
  ZOOM_LEVEL: "zoomLevel",
  SIDEBAR_COLLAPSED: "sidebarCollapsed"
};

export const SETTINGS = {
  GOOGLE_API_KEY: "googleApiKey",
  GOOGLE_OAUTH_CLIENT_ID: "googleOAuthClientId",
  ROLL_MODE: "rollMode", 
  DAMAGE_MODES: "damageModes",
  API_CACHE: "apiCache"
};

export const CSS_CLASSES = {
  SHEET: "sheexcel-sheet",
  CHECK_ENTRY: "sheexcel-check-entry",
  CHECK_ENTRY_SUB: "sheexcel-check-entry-sub",
  REFERENCE_ROW: "sheexcel-reference-row",
  ROLL_BUTTON: "sheexcel-roll",
  SIDEBAR: "sheexcel-sidebar",
  LOADING: "sheexcel-loading",
  ERROR: "sheexcel-error"
};

export const REFERENCE_TYPES = {
  CHECKS: "checks",
  SAVES: "saves", 
  ATTACKS: "attacks",
  SPELLS: "spells"
};

export const ERROR_MESSAGES = {
  NO_SHEET_URL: "No sheet URL set. Enter a Google Sheet URL in the settings sidebar and click Update Sheet.",
  INVALID_SHEET_URL: "Invalid Google Sheet URL format.",
  INVALID_CELL_REFERENCE: "Invalid cell reference format (e.g., A1, B2, etc.)",
  INVALID_SHEET_NAME: "Invalid sheet name.",
  API_ERROR: "Google Sheets API error",
  INVALID_BONUS: "Invalid bonus formula",
  NETWORK_ERROR: "Network error - check your connection"
};

export const API_CONFIG = {
  BASE_URL: "https://sheets.googleapis.com/v4/spreadsheets",
  CACHE_TTL: 5 * 60 * 1000, // 5 minutes
  TIMEOUT: 10000 // 10 seconds
};

export const VALIDATION = {
  CELL_REFERENCE_PATTERN: /^[A-Z]+[1-9]\d*(:[A-Z]+[1-9]\d*)?$/,
  SHEET_URL_PATTERN: /\/d\/([a-zA-Z0-9-_]+)/,
  SHEET_NAME_MAX_LENGTH: 100
};