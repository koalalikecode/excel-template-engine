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
 * Print area tracking
 */
export interface PrintAreaInfo {
    startRow: number;
    endRow: number;
    startCol: string;
    endCol: string;
    originalEndRow: number;
}

/**
 * Options for rendering Excel templates
 */
export interface RenderOptions {
    /**
     * When true, automatically converts numeric strings to numbers in Excel.
     * e.g. "100" → 100, "3.14" → 3.14
     * ⚠️ This may cause data loss for strings with leading zeros (e.g. "001234" → 1234)
     * @default false
     */
    autoParseNumbers?: boolean;
}
