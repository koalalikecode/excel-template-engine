import type { FormulaType } from './types';

// ============================================
// REGEX PATTERNS
// ============================================

/**
 * Regex patterns for placeholder detection
 */
export const PLACEHOLDER_REGEX = {
    /** Table anchor: {{#table path}} */
    TABLE: /^\{\{#table\s+(.+?)\}\}$/,

    /** Formula: {{#formula TYPE column/path}} */
    FORMULA: /^\{\{#formula\s+(SUM|AVERAGE|COUNT|MIN|MAX)\s+([A-Za-z_][\w.]*)\}\}$/i,

    /** Join array: {{path | join("separator")}} */
    JOIN: /^\{\{(.+?)\s*\|\s*join\(["'](.+?)["']\)\}\}$/,

    /** Simple value: {{path}} */
    VALUE: /^\{\{([^}]+)\}\}$/,

    /** Inline placeholder for interpolation */
    INLINE: /\{\{\s*([^}]+?)\s*\}\}/g,

    /** Column letter validation (A-ZZZ) */
    COLUMN_LETTER: /^[A-Z]+$/i,
} as const;

// ============================================
// FORMULA CONSTANTS
// ============================================

/**
 * Supported formula types
 */
export const FORMULA_TYPES: readonly FormulaType[] = [
    'SUM',
    'AVERAGE',
    'COUNT',
    'MIN',
    'MAX',
] as const;

// ============================================
// ERROR MESSAGES
// ============================================

/**
 * Error messages for the template engine
 */
export const ERROR_MESSAGES = {
    FIELD_NOT_FOUND: (field: string) => `Field '${field}' not found in table`,
    NO_COLUMN_SPECIFIED: 'No column specified for formula',
    NO_TABLE_OR_PATH: 'No table or array path found for formula',
    UNKNOWN_FORMULA_TYPE: (type: string) => `Unknown formula type: ${type}`,
    TEMPLATE_RENDER_FAILED: (msg: string) => `Failed to render template: ${msg}`,
} as const;
