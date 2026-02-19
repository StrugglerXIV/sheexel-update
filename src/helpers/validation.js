// helpers/validation.js
import { VALIDATION, ERROR_MESSAGES } from './constants.js';

export class ValidationError extends Error {
  constructor(message, field = null) {
    super(message);
    this.name = 'ValidationError';
    this.field = field;
  }
}

export function validateSheetUrl(url) {
  if (!url || typeof url !== 'string') {
    throw new ValidationError(ERROR_MESSAGES.NO_SHEET_URL);
  }

  const trimmed = url.trim();
  if (!trimmed) {
    throw new ValidationError(ERROR_MESSAGES.NO_SHEET_URL);
  }

  const match = trimmed.match(VALIDATION.SHEET_URL_PATTERN);
  if (!match) {
    throw new ValidationError(ERROR_MESSAGES.INVALID_SHEET_URL);
  }

  return {
    url: trimmed,
    sheetId: match[1]
  };
}

export function validateCellReference(cell) {
  if (!cell || typeof cell !== 'string') {
    throw new ValidationError(ERROR_MESSAGES.INVALID_CELL_REFERENCE, 'cell');
  }

  const trimmed = cell.trim().toUpperCase();
  if (!VALIDATION.CELL_REFERENCE_PATTERN.test(trimmed)) {
    throw new ValidationError(
      `${ERROR_MESSAGES.INVALID_CELL_REFERENCE}: "${cell}"`, 
      'cell'
    );
  }

  return trimmed;
}

export function validateSheetName(name) {
  if (!name || typeof name !== 'string') {
    throw new ValidationError(ERROR_MESSAGES.INVALID_SHEET_NAME, 'sheet');
  }

  const trimmed = name.trim();
  if (!trimmed || trimmed.length > VALIDATION.SHEET_NAME_MAX_LENGTH) {
    throw new ValidationError(ERROR_MESSAGES.INVALID_SHEET_NAME, 'sheet');
  }

  return trimmed;
}

export function validateReferenceData(ref) {
  const errors = [];

  try {
    ref.cell = validateCellReference(ref.cell);
  } catch (error) {
    errors.push(error.message);
  }

  try {
    ref.sheet = validateSheetName(ref.sheet);
  } catch (error) {
    errors.push(error.message);
  }

  if (!ref.keyword || typeof ref.keyword !== 'string') {
    errors.push("Keyword is required");
  }

  if (!ref.type || !['checks', 'saves', 'attacks', 'spells'].includes(ref.type)) {
    errors.push("Invalid reference type");
  }

  if (errors.length > 0) {
    throw new ValidationError(errors.join(', '));
  }

  return {
    ...ref,
    keyword: ref.keyword.trim()
  };
}

export function sanitizeSheetName(name) {
  const clean = validateSheetName(name);
  
  // Quote sheet names with special characters
  if (clean.match(/[^A-Za-z0-9_]/)) {
    return `'${clean.replace(/'/g, "''")}'`;
  }
  
  return clean;
}