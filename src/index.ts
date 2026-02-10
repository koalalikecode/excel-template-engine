import * as _ from 'lodash-es';
import ExcelJS from 'exceljs';
import { PLACEHOLDER_REGEX, ERROR_MESSAGES } from './constants';
import type {
    Placeholder,
    FormulaType,
    TableMetadata,
    TableRenderInfo,
    TableRenderResult,
    CellTemplate,
    MergePattern,
    PrintAreaInfo,
    RenderOptions,
} from './types';

// Re-export types for library consumers
export type { Placeholder, FormulaType, RenderOptions } from './types';

// ============================================
// 1. PATH RESOLUTION & SCOPE
// ============================================

/**
 * Resolve a path from a scope stack
 * Walks through scopes from most specific (local) to most general (root)
 */
function resolve(path: string, scopes: unknown[] = []): unknown {
    return _.reduce(
        scopes,
        (result: unknown, scope: unknown) => {
            if (!_.isUndefined(result)) return result;
            return _.get(scope as object, path);
        },
        undefined
    ) ?? "";
}

// ============================================
// 2. PLACEHOLDER DETECTION
// ============================================

/**
 * Detect and parse placeholder syntax in a cell value
 */
function detectPlaceholder(cellValue: unknown): Placeholder | null {
    if (!_.isString(cellValue)) return null;

    const trimmed = _.trim(cellValue);

    // Table anchor: {{#table path}}
    const tableMatch = trimmed.match(PLACEHOLDER_REGEX.TABLE);
    if (tableMatch) {
        return {
            type: 'table',
            path: _.trim(tableMatch[1])
        };
    }

    // Formula: {{#formula TYPE path}}
    const formulaMatch = trimmed.match(PLACEHOLDER_REGEX.FORMULA);
    if (formulaMatch) {
        const target = formulaMatch[2];
        const isColumnLetter = PLACEHOLDER_REGEX.COLUMN_LETTER.test(target) && target.length <= 3;

        if (isColumnLetter) {
            return {
                type: 'formula',
                path: '',
                formulaType: formulaMatch[1].toUpperCase() as FormulaType,
                column: target.toUpperCase()
            };
        } else {
            const lastDotIndex = target.lastIndexOf('.');
            const arrayPath = lastDotIndex > 0 ? target.substring(0, lastDotIndex) : '';
            const fieldName = lastDotIndex > 0 ? target.substring(lastDotIndex + 1) : target;

            return {
                type: 'formula',
                path: target,
                formulaType: formulaMatch[1].toUpperCase() as FormulaType,
                column: undefined,
                arrayPath,
                fieldName
            };
        }
    }

    // Join array: {{path | join(", ")}}
    const joinMatch = trimmed.match(PLACEHOLDER_REGEX.JOIN);
    if (joinMatch) {
        return {
            type: 'join',
            path: _.trim(joinMatch[1]),
            separator: joinMatch[2]
        };
    }

    // Simple value: {{path}}
    const valueMatch = trimmed.match(PLACEHOLDER_REGEX.VALUE);
    if (valueMatch) {
        return {
            type: 'value',
            path: _.trim(valueMatch[1])
        };
    }

    return null;
}

// ============================================
// 3. CELL RENDERING
// ============================================

/**
 * Render a value placeholder - preserves type for Excel
 */
function renderValue(path: string, scopes: unknown[], options: RenderOptions = {}): unknown {
    const value = resolve(path, scopes);

    if (_.isNil(value)) return "";
    if (_.isDate(value)) return value;
    if (_.isNumber(value)) return value;
    if (_.isString(value)) {
        const trimmed = value.trim();
        if (trimmed === '') return value;

        // Auto-parse numeric strings only when explicitly enabled
        if (options.autoParseNumbers && trimmed !== '' && !isNaN(Number(trimmed))) {
            return Number(trimmed);
        }

        // Only auto-parse strict ISO date strings (e.g. "2024-01-15", "2024-01-15T10:30:00Z")
        if (/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?)?$/.test(trimmed)) {
            const timestamp = Date.parse(trimmed);
            if (!isNaN(timestamp)) {
                return new Date(timestamp);
            }
        }

        return value;
    }
    if (_.isBoolean(value)) return value;

    return JSON.stringify(value);
}

/**
 * Render a formula placeholder to Excel formula object or calculated value
 */
function renderFormula(
    formulaType: FormulaType,
    columnOrField: string | undefined,
    fieldName: string | undefined,
    arrayPath: string | undefined,
    tableInfo: TableRenderInfo | null,
    data: unknown
): { formula: string } | number | string {
    if (tableInfo && tableInfo.startRow > 0 && tableInfo.endRow > 0) {
        let column: string | undefined = columnOrField;

        if (!column && fieldName) {
            column = tableInfo.fieldColumnMap.get(fieldName);
            if (!column) {
                return `#ERROR: ${ERROR_MESSAGES.FIELD_NOT_FOUND(fieldName)}`;
            }
        }

        if (!column) {
            return `#ERROR: ${ERROR_MESSAGES.NO_COLUMN_SPECIFIED}`;
        }

        const formula = `${formulaType}(${column}${tableInfo.startRow}:${column}${tableInfo.endRow})`;
        return { formula };
    }

    if (!arrayPath || !fieldName) {
        return `#ERROR: ${ERROR_MESSAGES.NO_TABLE_OR_PATH}`;
    }

    const arr = _.get(data as object, arrayPath);
    if (!_.isArray(arr) || _.isEmpty(arr)) {
        return 0;
    }

    const values = arr
        .map(item => {
            const val = _.get(item as object, fieldName);
            if (_.isNumber(val)) return val;
            if (_.isString(val) && val.trim() !== '' && !isNaN(Number(val))) {
                return Number(val);
            }
            return NaN;
        })
        .filter(v => !isNaN(v));

    if (values.length === 0) return 0;

    switch (formulaType) {
        case 'SUM':
            return _.sum(values);
        case 'AVERAGE':
            return _.mean(values);
        case 'COUNT':
            return values.length;
        case 'MIN':
            return _.min(values) ?? 0;
        case 'MAX':
            return _.max(values) ?? 0;
        default:
            return `#ERROR: ${ERROR_MESSAGES.UNKNOWN_FORMULA_TYPE(formulaType)}`;
    }
}

/**
 * Join an array into a single cell value
 */
function renderJoin(path: string, separator: string, scopes: unknown[]): string {
    const lastDotIndex = path.lastIndexOf('.');

    if (lastDotIndex > 0) {
        const arrayPath = path.substring(0, lastDotIndex);
        const propertyName = path.substring(lastDotIndex + 1);
        const arr = resolve(arrayPath, scopes);

        if (_.isArray(arr) && !_.isEmpty(arr)) {
            return _.chain(arr)
                .map(item => {
                    const val = _.get(item as object, propertyName);
                    return _.isNil(val) ? "" : String(val);
                })
                .filter(v => v !== "")
                .join(separator)
                .value();
        }
    }

    const arr = resolve(path, scopes);

    if (!_.isArray(arr) || _.isEmpty(arr)) return "";

    return _.chain(arr)
        .map(item => _.isNil(item) ? "" : String(item))
        .filter(v => v !== "" && v !== "[object Object]")
        .join(separator)
        .value();
}

/**
 * Calculate formula value directly from data (for interpolated strings)
 */
function calculateFormulaValue(
    formulaType: string,
    arrayPath: string,
    fieldName: string,
    scopes: unknown[]
): string {
    const arr = resolve(arrayPath, scopes);
    if (!_.isArray(arr) || _.isEmpty(arr)) {
        return "0";
    }

    const values = arr
        .map(item => {
            const val = _.get(item as object, fieldName);
            if (_.isNumber(val)) return val;
            if (_.isString(val) && val.trim() !== '' && !isNaN(Number(val))) {
                return Number(val);
            }
            return NaN;
        })
        .filter(v => !isNaN(v));

    if (values.length === 0) return "0";

    let result: number;
    switch (formulaType.toUpperCase()) {
        case 'SUM':
            result = _.sum(values);
            break;
        case 'AVERAGE':
            result = _.mean(values);
            break;
        case 'COUNT':
            result = values.length;
            break;
        case 'MIN':
            result = _.min(values) ?? 0;
            break;
        case 'MAX':
            result = _.max(values) ?? 0;
            break;
        default:
            return `#ERROR: Unknown formula type: ${formulaType}`;
    }

    return String(result);
}

/**
 * Render an interpolated string with multiple placeholders
 * Supports: {{value}}, {{path | join("sep")}}, {{#formula TYPE path}}
 */
function renderInterpolatedString(template: string, scopes: unknown[]): string {
    return template.replace(PLACEHOLDER_REGEX.INLINE, (fullMatch, exprContent) => {
        const expr = _.trim(exprContent);

        // Handle formula: {{#formula SUM items.amount}}
        const formulaMatch = expr.match(/^#formula\s+(SUM|AVERAGE|COUNT|MIN|MAX)\s+([A-Za-z_][\w.]*)$/i);
        if (formulaMatch) {
            const formulaType = formulaMatch[1];
            const target = formulaMatch[2];
            const lastDotIndex = target.lastIndexOf('.');

            if (lastDotIndex > 0) {
                const arrayPath = target.substring(0, lastDotIndex);
                const fieldName = target.substring(lastDotIndex + 1);
                return calculateFormulaValue(formulaType, arrayPath, fieldName, scopes);
            }
            return `#ERROR: Invalid formula path: ${target}`;
        }

        // Handle join: {{items.name | join(", ")}}
        const joinMatch = expr.match(/^(.+?)\s*\|\s*join\(["'](.+?)["']\)$/);
        if (joinMatch) {
            const path = _.trim(joinMatch[1]);
            const separator = joinMatch[2];
            return renderJoin(path, separator, scopes);
        }

        // Handle simple value: {{name}}
        const value = resolve(expr, scopes);
        if (_.isNil(value)) return "";
        return String(value);
    });
}

function handleRichText(cellValue: unknown, scopes: unknown[]): unknown {
    if (!cellValue || !Array.isArray((cellValue as { richText?: unknown[] }).richText)) {
        return cellValue;
    }

    const richTextValue = cellValue as { richText: Array<{ text: string; font?: unknown }> };
    const result: Array<{ text: string; font?: unknown }> = [];
    let buffer: Array<{ text: string; font?: unknown }> = [];
    let inPlaceholder = false;

    for (const fragment of richTextValue.richText) {
        const { text } = fragment;

        if (text.includes('{{') || inPlaceholder) {
            inPlaceholder = true;
            buffer.push(fragment);

            if (text.includes('}}')) {
                const combinedText = buffer.map(f => f.text).join('');
                const rendered = renderInterpolatedString(combinedText, scopes);

                result.push({
                    text: rendered,
                    font: buffer[0]?.font
                });

                buffer = [];
                inPlaceholder = false;
            }
            continue;
        }

        result.push(fragment);
    }

    if (buffer.length > 0) {
        result.push(...buffer);
    }

    return { richText: result };
}

/**
 * Render a cell based on its placeholder type or content
 */
function renderCell(cellValue: unknown, scopes: unknown[], options: RenderOptions = {}): unknown {
    const placeholder = detectPlaceholder(cellValue);

    if (placeholder) {
        switch (placeholder.type) {
            case 'value':
                return renderValue(placeholder.path, scopes, options);
            case 'join':
                return renderJoin(placeholder.path, placeholder.separator!, scopes);
            case 'table':
                return cellValue;
        }
    }

    if (_.isString(cellValue) && cellValue.includes('{{')) {
        return renderInterpolatedString(cellValue, scopes);
    }

    if (_.isObject(cellValue) && 'richText' in cellValue) {
        return handleRichText(cellValue, scopes);
    }

    return cellValue;
}

// ============================================
// 4. PRINT AREA & PAGE BREAKS MANAGEMENT
// ============================================

function parsePrintArea(printArea: string | undefined): PrintAreaInfo | null {
    if (!printArea) return null;

    const areaStr = printArea.includes('!') ? printArea.split('!')[1] : printArea;
    const match = areaStr.match(/\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)/);
    if (!match) return null;

    return {
        startCol: match[1],
        startRow: parseInt(match[2]),
        endCol: match[3],
        endRow: parseInt(match[4]),
        originalEndRow: parseInt(match[4])
    };
}

function adjustImageAnchors(
    worksheet: ExcelJS.Worksheet,
    insertPosition: number,
    rowsAdded: number
): void {
    const wsAny = worksheet as unknown as { getImages?: () => Array<{ range?: { tl?: { row: number; nativeRow: number }; br?: { row: number; nativeRow: number } }; position?: { row: number } }>; model?: { media?: Array<{ range?: { tl?: { nativeRow: number }; br?: { nativeRow: number } } }> } };
    const images = wsAny.getImages?.() || [];

    if (images.length === 0) {
        const model = wsAny.model;
        if (model?.media) {
            model.media.forEach((image) => {
                if (image.range) {
                    if (image.range.tl && image.range.tl.nativeRow >= insertPosition - 1) {
                        image.range.tl.nativeRow += rowsAdded;
                    }
                    if (image.range.br && image.range.br.nativeRow >= insertPosition - 1) {
                        image.range.br.nativeRow += rowsAdded;
                    }
                }
            });
        }
        return;
    }

    images.forEach((image) => {
        if (image.range) {
            if (image.range.tl && image.range.tl.row >= insertPosition - 1) {
                image.range.tl.row += rowsAdded;
            }
            if (image.range.br && image.range.br.row >= insertPosition - 1) {
                image.range.br.row += rowsAdded;
            }
        } else if (image.position) {
            if (image.position.row >= insertPosition - 1) {
                image.position.row += rowsAdded;
            }
        }
    });
}

function recalculatePrintArea(worksheet: ExcelJS.Worksheet): void {
    let minRow = Infinity;
    let maxRow = 0;
    let minCol = Infinity;
    let maxCol = 0;
    let hasContent = false;

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            if (cell.value !== null && cell.value !== undefined && cell.value !== '') {
                hasContent = true;
                minRow = Math.min(minRow, rowNumber);
                maxRow = Math.max(maxRow, rowNumber);
                minCol = Math.min(minCol, colNumber);
                maxCol = Math.max(maxCol, colNumber);
            }
        });
    });

    if (!hasContent) {
        worksheet.pageSetup.printArea = undefined;
        return;
    }

    const startCol = columnNumberToLetter(minCol);
    const endCol = columnNumberToLetter(maxCol);
    worksheet.pageSetup.printArea = `${startCol}${minRow}:${endCol}${maxRow}`;
}

// ============================================
// 5. TABLE RENDERING UTILITIES
// ============================================

function columnLetterToNumber(letter: string): number {
    const upperLetter = letter.toUpperCase();
    let num = 0;
    for (let i = 0; i < upperLetter.length; i++) {
        num = num * 26 + (upperLetter.charCodeAt(i) - 64);
    }
    return num;
}

function columnNumberToLetter(num: number): string {
    let letter = '';
    while (num > 0) {
        const mod = (num - 1) % 26;
        letter = String.fromCharCode(65 + mod) + letter;
        num = Math.floor((num - mod) / 26);
    }
    return letter;
}

function getMergeInfo(
    worksheet: ExcelJS.Worksheet,
    rowNumber: number,
    colNumber: number
): { master: { row: number; col: number }; range: string } | null {
    const wsModel = worksheet.model as { merges?: Array<string | { top: number; left: number; bottom: number; right: number }> };
    const merges = wsModel.merges || [];

    for (const merge of merges) {
        let top: number; let left: number; let bottom: number; let right: number;

        if (typeof merge === 'string') {
            const [start, end] = merge.split(':');
            const startMatch = start.match(/([A-Z]+)(\d+)/);
            const endMatch = end.match(/([A-Z]+)(\d+)/);

            if (!startMatch || !endMatch) continue;

            left = columnLetterToNumber(startMatch[1]);
            top = parseInt(startMatch[2]);
            right = columnLetterToNumber(endMatch[1]);
            bottom = parseInt(endMatch[2]);
        } else {
            top = merge.top;
            left = merge.left;
            bottom = merge.bottom;
            right = merge.right;
        }

        if (rowNumber >= top && rowNumber <= bottom &&
            colNumber >= left && colNumber <= right) {
            return {
                master: { row: top, col: left },
                range: `${columnNumberToLetter(left)}${top}:${columnNumberToLetter(right)}${bottom}`
            };
        }
    }

    return null;
}

function extractMergePatternsFromRow(
    worksheet: ExcelJS.Worksheet,
    rowNumber: number
): MergePattern[] {
    const patterns: MergePattern[] = [];
    const processedCols = new Set<number>();

    const row = worksheet.getRow(rowNumber);
    const maxCol = row.cellCount || 20;

    for (let col = 1; col <= maxCol; col++) {
        if (processedCols.has(col)) continue;

        const mergeInfo = getMergeInfo(worksheet, rowNumber, col);

        if (mergeInfo && mergeInfo.master.row === rowNumber) {
            const [start, end] = mergeInfo.range.split(':');
            const startColMatch = start.match(/[A-Z]+/i);
            const endColMatch = end.match(/[A-Z]+/i);

            if (!startColMatch || !endColMatch) continue;

            const startCol = columnLetterToNumber(startColMatch[0]);
            const endCol = columnLetterToNumber(endColMatch[0]);

            patterns.push({
                startCol,
                endCol,
                spanCols: endCol - startCol + 1
            });

            for (let c = startCol; c <= endCol; c++) {
                processedCols.add(c);
            }
        }
    }

    return patterns;
}

function extractRowTemplate(
    row: ExcelJS.Row,
    worksheet: ExcelJS.Worksheet
): { templates: CellTemplate[]; mergePatterns: MergePattern[]; fieldColumnMap: Map<string, string> } {
    const templates: CellTemplate[] = [];
    const fieldColumnMap: Map<string, string> = new Map();
    const rowNumber = row.number;

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const mergeInfo = getMergeInfo(worksheet, rowNumber, colNumber);

        templates[colNumber] = {
            value: cell.value,
            style: _.cloneDeep(cell.style),
            numFmt: cell.numFmt,
            type: cell.type,
            isMerged: !!mergeInfo,
            mergeInfo: mergeInfo ? {
                startCol: mergeInfo.master.col,
                endCol: (() => {
                    const endMatch = mergeInfo.range.split(':')[1].match(/[A-Z]+/i);
                    return endMatch ? columnLetterToNumber(endMatch[0]) : mergeInfo.master.col;
                })()
            } : undefined
        };

        const cellValue = _.isString(cell.value) ? cell.value : '';
        const placeholderRegex = /\{\{(\w+(?:\.\w+)*)\}\}/g;
        let match;
        while ((match = placeholderRegex.exec(cellValue)) !== null) {
            const fieldPath = match[1];
            const columnLetter = columnNumberToLetter(colNumber);
            fieldColumnMap.set(fieldPath, columnLetter);

            const lastSegment = fieldPath.split('.').pop();
            if (lastSegment && !fieldColumnMap.has(lastSegment)) {
                fieldColumnMap.set(lastSegment, columnLetter);
            }
        }
    });

    const mergePatterns = extractMergePatternsFromRow(worksheet, rowNumber);

    return { templates, mergePatterns, fieldColumnMap };
}

function applyMergePatternsToRow(
    worksheet: ExcelJS.Worksheet,
    rowNumber: number,
    patterns: MergePattern[]
): void {
    patterns.forEach(pattern => {
        const startCell = `${columnNumberToLetter(pattern.startCol)}${rowNumber}`;
        const endCell = `${columnNumberToLetter(pattern.endCol)}${rowNumber}`;
        const range = `${startCell}:${endCell}`;

        try {
            worksheet.unMergeCells(range);
            worksheet.mergeCells(range);
        } catch {
            // Ignore merge errors
        }
    });
}

function applyRowTemplate(
    newRow: ExcelJS.Row,
    templates: CellTemplate[],
    scopes: unknown[],
    options: RenderOptions = {}
): void {
    const processedCols = new Set<number>();

    _.forEach(templates, (cellTemplate, colNumber) => {
        if (!cellTemplate || processedCols.has(colNumber)) return;

        const cell = newRow.getCell(colNumber);

        if (cellTemplate.isMerged && cellTemplate.mergeInfo) {
            const { startCol, endCol } = cellTemplate.mergeInfo;

            if (colNumber === startCol) {
                const renderedValue = renderCell(cellTemplate.value, scopes, options);
                cell.value = renderedValue as ExcelJS.CellValue;
                cell.style = cellTemplate.style;
                if (cellTemplate.numFmt) cell.numFmt = cellTemplate.numFmt;

                for (let c = startCol; c <= endCol; c++) {
                    processedCols.add(c);
                }
            }
        } else {
            const renderedValue = renderCell(cellTemplate.value, scopes, options);
            cell.value = renderedValue as ExcelJS.CellValue;
            cell.style = cellTemplate.style;
            if (cellTemplate.numFmt) cell.numFmt = cellTemplate.numFmt;
        }
    });
}

function clearAndHideRow(row: ExcelJS.Row): void {
    row.eachCell({ includeEmpty: true }, (cell) => {
        cell.value = null;
    });
    row.hidden = true;
    row.height = 0;
}

function calculateHideOffset(
    entryIndex: number,
    hideEntries: { rows: number[], insertedAt: number, rowsInserted: number }[]
): number {
    let offset = 0;
    for (let j = entryIndex + 1; j < hideEntries.length; j++) {
        const laterEntry = hideEntries[j];
        const currentEntry = hideEntries[entryIndex];
        if (laterEntry.insertedAt < currentEntry.insertedAt && laterEntry.rowsInserted > 0) {
            offset += laterEntry.rowsInserted;
        }
    }
    return offset;
}

function renderTable(
    worksheet: ExcelJS.Worksheet,
    anchorRowIndex: number,
    arrayPath: string,
    scopes: unknown[],
    rootData: unknown,
    _printAreaInfo: PrintAreaInfo | null,
    options: RenderOptions = {}
): TableRenderResult {
    const allScopes = [...scopes, rootData];
    const arr = resolve(arrayPath, allScopes);

    if (!_.isArray(arr)) {
        return {
            tableInfo: null,
            rowsToHide: [anchorRowIndex, anchorRowIndex + 1]
        };
    }

    const templateRowIndex = anchorRowIndex + 1;
    const templateRow = worksheet.getRow(templateRowIndex);
    const { templates, mergePatterns, fieldColumnMap } = extractRowTemplate(templateRow, worksheet);

    if (_.isEmpty(arr)) {
        return {
            tableInfo: null,
            rowsToHide: [anchorRowIndex, anchorRowIndex + 1]
        };
    }

    const rowsToAdd = arr.length;
    const startRow = anchorRowIndex;
    const endRow = anchorRowIndex + rowsToAdd - 1;

    _.forEach(arr, (item, index) => {
        const insertPosition = anchorRowIndex + index;
        const newRow = worksheet.insertRow(insertPosition, []);
        const localScopes = [item, ...allScopes];

        applyRowTemplate(newRow, templates, localScopes, options);
        applyMergePatternsToRow(worksheet, insertPosition, mergePatterns);
    });

    const newAnchorRowIndex = anchorRowIndex + rowsToAdd;
    const newTemplateRowIndex = anchorRowIndex + rowsToAdd + 1;

    adjustImageAnchors(worksheet, anchorRowIndex, rowsToAdd);

    return {
        tableInfo: {
            startRow,
            endRow,
            path: arrayPath,
            fieldColumnMap
        },
        rowsToHide: [newAnchorRowIndex, newTemplateRowIndex]
    };
}

// ============================================
// 6. SCAN & COLLECT TABLES
// ============================================

function scanTables(worksheet: ExcelJS.Worksheet): TableMetadata[] {
    const tables: TableMetadata[] = [];
    const seenAnchors: Set<string> = new Set();

    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell) => {
            const placeholder = detectPlaceholder(cell.value);

            if (placeholder?.type === 'table') {
                const anchorKey = `${rowNumber}:${placeholder.path}`;

                if (seenAnchors.has(anchorKey)) {
                    return;
                }

                seenAnchors.add(anchorKey);
                tables.push({
                    rowNumber,
                    path: placeholder.path
                });
            }
        });
    });

    return _.orderBy(tables, ['rowNumber'], ['desc']);
}

// ============================================
// 7. MAIN ENGINE
// ============================================

function processWorkbook(workbook: ExcelJS.Workbook, data: unknown, options: RenderOptions = {}): void {
    _.forEach(workbook.worksheets, (worksheet) => {
        const columnWidths: Map<number, number | undefined> = new Map();
        worksheet.columns.forEach((col, index) => {
            if (col.width) {
                columnWidths.set(index + 1, col.width);
            }
        });

        const originallyHiddenRows: Set<number> = new Set();
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            if (row.hidden) {
                originallyHiddenRows.add(rowNumber);
            }
        });

        const printAreaInfo = parsePrintArea(worksheet.pageSetup.printArea as string);
        const tables = scanTables(worksheet);
        const tableInfoMap: Map<string, TableRenderInfo> = new Map();
        let lastTableInfo: TableRenderInfo | null = null;

        const hideEntries: { rows: number[], insertedAt: number, rowsInserted: number }[] = [];

        _.forEach(tables, ({ rowNumber, path }) => {
            const result = renderTable(worksheet, rowNumber, path, [], data, printAreaInfo, options);

            if (result.tableInfo) {
                tableInfoMap.set(path, result.tableInfo);
                lastTableInfo = result.tableInfo;
            }

            const rowsInserted = result.tableInfo ? (result.tableInfo.endRow - result.tableInfo.startRow + 1) : 0;
            hideEntries.push({
                rows: result.rowsToHide,
                insertedAt: rowNumber,
                rowsInserted: rowsInserted
            });
        });

        const insertInfo: { insertedAt: number, count: number }[] = hideEntries
            .filter(e => e.rowsInserted > 0)
            .map(e => ({ insertedAt: e.insertedAt, count: e.rowsInserted }));

        for (let i = 0; i < hideEntries.length; i++) {
            const entry = hideEntries[i];
            const offset = calculateHideOffset(i, hideEntries);

            entry.rows.forEach(rowIndex => {
                clearAndHideRow(worksheet.getRow(rowIndex + offset));
            });
        }

        const anchorTemplateRows = new Set<number>();
        hideEntries.forEach((entry, i) => {
            const offset = calculateHideOffset(i, hideEntries);
            entry.rows.forEach(r => anchorTemplateRows.add(r + offset));
        });

        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            if (row.hidden && !anchorTemplateRows.has(rowNumber)) {
                row.hidden = false;
            }
        });

        originallyHiddenRows.forEach(originalRowNum => {
            let shiftAmount = 0;
            insertInfo.forEach(info => {
                if (info.insertedAt <= originalRowNum) {
                    shiftAmount += info.count;
                }
            });

            const newRowNum = originalRowNum + shiftAmount;
            const isTableRow = tables.some(t =>
                originalRowNum === t.rowNumber || originalRowNum === t.rowNumber + 1
            );

            if (!isTableRow) {
                const row = worksheet.getRow(newRowNum);
                row.hidden = true;
            }
        });

        worksheet.eachRow((row) => {
            row.eachCell((cell) => {
                const placeholder = detectPlaceholder(cell.value);
                if (!placeholder) {
                    cell.value = renderCell(cell.value, [data], options) as ExcelJS.CellValue;
                    return;
                }

                if (placeholder.type === 'table') return;

                if (placeholder.type === 'formula' && placeholder.formulaType) {
                    let tableInfo: TableRenderInfo | null = null;

                    if (placeholder.column) {
                        tableInfo = lastTableInfo;
                    } else if (placeholder.arrayPath) {
                        tableInfo = tableInfoMap.get(placeholder.arrayPath) || null;
                    } else if (placeholder.fieldName) {
                        tableInfo = lastTableInfo;
                    }

                    cell.value = renderFormula(
                        placeholder.formulaType,
                        placeholder.column,
                        placeholder.fieldName,
                        placeholder.arrayPath,
                        tableInfo,
                        data
                    ) as ExcelJS.CellValue;
                    return;
                }

                cell.value = renderCell(cell.value, [data], options) as ExcelJS.CellValue;
            });
        });

        recalculatePrintArea(worksheet);

        columnWidths.forEach((width, colNumber) => {
            if (width) {
                const col = worksheet.getColumn(colNumber);
                col.width = width;
            }
        });
    });
}

/**
 * Main rendering function - processes Excel template with data
 */
async function renderExcelTemplate(
    templatePath: string,
    data: unknown,
    outputPath: string,
    options: RenderOptions = {}
): Promise<void> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    processWorkbook(workbook, data, options);

    await workbook.xlsx.writeFile(outputPath);
}

/**
 * Render Excel template from buffer
 */
async function renderExcelTemplateFromBuffer(
    templateBuffer: ArrayBuffer | Buffer | Uint8Array,
    data: unknown,
    options: RenderOptions = {}
): Promise<Buffer> {
    const workbook = new ExcelJS.Workbook();

    let bufferToLoad: ArrayBuffer;

    if (templateBuffer instanceof ArrayBuffer) {
        bufferToLoad = templateBuffer;
    } else {
        bufferToLoad = templateBuffer.buffer.slice(
            templateBuffer.byteOffset,
            templateBuffer.byteOffset + templateBuffer.byteLength
        ) as ArrayBuffer;
    }

    await workbook.xlsx.load(bufferToLoad);
    processWorkbook(workbook, data, options);

    const outputBuffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(outputBuffer);
}

// ============================================
// 8. EXPORTS
// ============================================

export {
    // Main functions
    renderExcelTemplate,
    renderExcelTemplateFromBuffer,
};