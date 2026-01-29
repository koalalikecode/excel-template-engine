import type ExcelJS from 'exceljs';

// ============================================
// PLACEHOLDER TYPES
// ============================================

/**
 * Placeholder types supported by the engine
 */
export type PlaceholderType = 'value' | 'table' | 'join' | 'formula';

/**
 * Formula types supported
 */
export type FormulaType = 'SUM' | 'AVERAGE' | 'COUNT' | 'MIN' | 'MAX';

// ============================================
// PLACEHOLDER INTERFACES
// ============================================

/**
 * Detected placeholder structure
 */
export interface Placeholder {
    type: PlaceholderType;
    path: string;
    separator?: string;
    formulaType?: FormulaType;
    column?: string;
    arrayPath?: string;
    fieldName?: string;
}

// ============================================
// TABLE INTERFACES
// ============================================

/**
 * Table metadata found during scanning
 */
export interface TableMetadata {
    rowNumber: number;
    path: string;
}

/**
 * Table render info for formula generation
 */
export interface TableRenderInfo {
    startRow: number;
    endRow: number;
    path: string;
    fieldColumnMap: Map<string, string>;
}

/**
 * Table render result
 */
export interface TableRenderResult {
    tableInfo: TableRenderInfo | null;
    rowsToHide: number[];
}

// ============================================
// CELL INTERFACES
// ============================================

/**
 * Cell template structure for cloning rows
 */
export interface CellTemplate {
    value: ExcelJS.CellValue;
    style: Partial<ExcelJS.Style>;
    numFmt?: string;
    type?: ExcelJS.ValueType;
    isMerged?: boolean;
    mergeInfo?: { startCol: number; endCol: number };
}

/**
 * Merge pattern for row templates
 */
export interface MergePattern {
    startCol: number;
    endCol: number;
    spanCols: number;
}

// ============================================
// UTILITY INTERFACES
// ============================================

/**
 * Data validation result
 */
export interface ValidationResult {
    valid: boolean;
    missing: string[];
}

/**
 * Debug placeholder information
 */
export interface PlaceholderDebugInfo {
    sheet: number;
    cell: string;
    type: PlaceholderType;
    path: string;
    separator?: string;
}

/**
 * Print area tracking
 */
export interface PrintAreaInfo {
    startRow: number;
    endRow: number;
    startCol: string;
    endCol: string;
    originalEndRow: number;
}
