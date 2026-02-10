import { describe, it, expect } from 'vitest';
import {
    _resolve,
    _detectPlaceholder,
    _renderValue,
    _renderFormula,
    _renderJoin,
    _renderInterpolatedString,
    _renderCell,
    _calculateFormulaValue,
    renderExcelTemplateFromBuffer,
} from './index';
import ExcelJS from 'exceljs';

// ============================================
// 1. resolve
// ============================================

describe('resolve', () => {
    it('should resolve simple path from scope', () => {
        const scopes = [{ name: 'John' }];
        expect(_resolve('name', scopes)).toBe('John');
    });

    it('should resolve nested path', () => {
        const scopes = [{ company: { name: 'Acme' } }];
        expect(_resolve('company.name', scopes)).toBe('Acme');
    });

    it('should return "" for missing path', () => {
        const scopes = [{ name: 'John' }];
        expect(_resolve('age', scopes)).toBe('');
    });

    it('should resolve from most specific scope first', () => {
        const scopes = [
            { name: 'Local' },
            { name: 'Root' },
        ];
        expect(_resolve('name', scopes)).toBe('Local');
    });

    it('should fall back to broader scope', () => {
        const scopes = [
            { age: 25 },
            { name: 'Root' },
        ];
        expect(_resolve('name', scopes)).toBe('Root');
    });

    it('should return "" for empty scopes', () => {
        expect(_resolve('name', [])).toBe('');
    });
});

// ============================================
// 2. detectPlaceholder
// ============================================

describe('detectPlaceholder', () => {
    it('should detect simple value placeholder', () => {
        const result = _detectPlaceholder('{{name}}');
        expect(result).toEqual({ type: 'value', path: 'name' });
    });

    it('should detect nested value placeholder', () => {
        const result = _detectPlaceholder('{{company.name}}');
        expect(result).toEqual({ type: 'value', path: 'company.name' });
    });

    it('should detect table placeholder', () => {
        const result = _detectPlaceholder('{{#table items}}');
        expect(result).toEqual({ type: 'table', path: 'items' });
    });

    it('should detect join placeholder', () => {
        const result = _detectPlaceholder('{{tags | join(", ")}}');
        expect(result).toEqual({
            type: 'join',
            path: 'tags',
            separator: ', ',
        });
    });

    it('should detect formula with column letter', () => {
        const result = _detectPlaceholder('{{#formula SUM C}}');
        expect(result).toEqual({
            type: 'formula',
            path: '',
            formulaType: 'SUM',
            column: 'C',
        });
    });

    it('should detect formula with field path', () => {
        const result = _detectPlaceholder('{{#formula SUM items.price}}');
        expect(result).toEqual({
            type: 'formula',
            path: 'items.price',
            formulaType: 'SUM',
            column: undefined,
            arrayPath: 'items',
            fieldName: 'price',
        });
    });

    it('should detect formula types case-insensitively', () => {
        const result = _detectPlaceholder('{{#formula average C}}');
        expect(result?.formulaType).toBe('AVERAGE');
    });

    it('should return null for non-placeholder strings', () => {
        expect(_detectPlaceholder('hello world')).toBeNull();
    });

    it('should return null for non-string values', () => {
        expect(_detectPlaceholder(123)).toBeNull();
        expect(_detectPlaceholder(null)).toBeNull();
        expect(_detectPlaceholder(undefined)).toBeNull();
    });
});

// ============================================
// 3. renderValue
// ============================================

describe('renderValue', () => {
    it('should return string as-is (no auto number conversion)', () => {
        const scopes = [{ code: '001234' }];
        expect(_renderValue('code', scopes)).toBe('001234');
    });

    it('should NOT auto-convert numeric string to number by default', () => {
        const scopes = [{ price: '100' }];
        expect(_renderValue('price', scopes)).toBe('100');
    });

    it('should auto-convert numeric string when autoParseNumbers is true', () => {
        const scopes = [{ price: '100' }];
        expect(_renderValue('price', scopes, { autoParseNumbers: true })).toBe(100);
    });

    it('should preserve leading zeros when autoParseNumbers is false', () => {
        const scopes = [{ zip: '00501' }];
        expect(_renderValue('zip', scopes)).toBe('00501');
    });

    it('should preserve number type', () => {
        const scopes = [{ amount: 42 }];
        expect(_renderValue('amount', scopes)).toBe(42);
    });

    it('should preserve Date type', () => {
        const date = new Date('2024-01-15');
        const scopes = [{ date }];
        expect(_renderValue('date', scopes)).toEqual(date);
    });

    it('should auto-parse strict ISO date string', () => {
        const scopes = [{ date: '2024-01-15' }];
        const result = _renderValue('date', scopes);
        expect(result).toBeInstanceOf(Date);
    });

    it('should auto-parse ISO datetime string', () => {
        const scopes = [{ date: '2024-01-15T10:30:00Z' }];
        const result = _renderValue('date', scopes);
        expect(result).toBeInstanceOf(Date);
    });

    it('should NOT auto-parse non-strict date-like strings', () => {
        const scopes = [{ code: '2024-01-15-report' }];
        expect(_renderValue('code', scopes)).toBe('2024-01-15-report');
    });

    it('should return "" for null/undefined', () => {
        const scopes = [{ name: null }];
        expect(_renderValue('name', scopes)).toBe('');
        expect(_renderValue('missing', scopes)).toBe('');
    });

    it('should preserve boolean values', () => {
        const scopes = [{ active: true }];
        expect(_renderValue('active', scopes)).toBe(true);
    });

    it('should JSON.stringify objects', () => {
        const scopes = [{ meta: { a: 1 } }];
        expect(_renderValue('meta', scopes)).toBe('{"a":1}');
    });
});

// ============================================
// 4. renderFormula
// ============================================

describe('renderFormula', () => {
    it('should generate SUM formula with table info', () => {
        const tableInfo = {
            startRow: 5,
            endRow: 10,
            path: 'items',
            fieldColumnMap: new Map([['price', 'C']]),
        };
        const result = _renderFormula('SUM', undefined, 'price', 'items', tableInfo, {});
        expect(result).toEqual({ formula: 'SUM(C5:C10)' });
    });

    it('should generate formula with column letter', () => {
        const tableInfo = {
            startRow: 3,
            endRow: 7,
            path: 'items',
            fieldColumnMap: new Map(),
        };
        const result = _renderFormula('AVERAGE', 'D', undefined, undefined, tableInfo, {});
        expect(result).toEqual({ formula: 'AVERAGE(D3:D7)' });
    });

    it('should calculate fallback value when no table', () => {
        const data = { items: [{ price: 10 }, { price: 20 }, { price: 30 }] };
        const result = _renderFormula('SUM', undefined, 'price', 'items', null, data);
        expect(result).toBe(60);
    });

    it('should return 0 for empty array fallback', () => {
        const data = { items: [] };
        const result = _renderFormula('SUM', undefined, 'price', 'items', null, data);
        expect(result).toBe(0);
    });

    it('should return error for unknown formula type', () => {
        const data = { items: [{ price: 10 }] };
        const result = _renderFormula('MEDIAN' as any, undefined, 'price', 'items', null, data);
        expect(result).toContain('#ERROR');
    });
});

// ============================================
// 5. renderJoin
// ============================================

describe('renderJoin', () => {
    it('should join primitive array', () => {
        const scopes = [{ tags: ['a', 'b', 'c'] }];
        expect(_renderJoin('tags', ', ', scopes)).toBe('a, b, c');
    });

    it('should join object array property', () => {
        const scopes = [{
            items: [{ name: 'Apple' }, { name: 'Banana' }]
        }];
        expect(_renderJoin('items.name', ' / ', scopes)).toBe('Apple / Banana');
    });

    it('should return "" for empty array', () => {
        const scopes = [{ tags: [] }];
        expect(_renderJoin('tags', ', ', scopes)).toBe('');
    });

    it('should return "" for missing path', () => {
        const scopes = [{}];
        expect(_renderJoin('tags', ', ', scopes)).toBe('');
    });

    it('should filter out nil values', () => {
        const scopes = [{
            items: [{ name: 'Apple' }, { name: null }, { name: 'Cherry' }]
        }];
        expect(_renderJoin('items.name', ', ', scopes)).toBe('Apple, Cherry');
    });
});

// ============================================
// 6. renderInterpolatedString
// ============================================

describe('renderInterpolatedString', () => {
    it('should interpolate simple values', () => {
        const scopes = [{ name: 'John', city: 'NY' }];
        expect(_renderInterpolatedString('Hello {{name}} from {{city}}', scopes))
            .toBe('Hello John from NY');
    });

    it('should handle missing values as empty string', () => {
        const scopes = [{ name: 'John' }];
        expect(_renderInterpolatedString('Hello {{name}} {{missing}}', scopes))
            .toBe('Hello John ');
    });

    it('should handle inline join', () => {
        const scopes = [{ tags: ['a', 'b'] }];
        expect(_renderInterpolatedString('Tags: {{tags | join(", ")}}', scopes))
            .toBe('Tags: a, b');
    });

    it('should handle inline formula', () => {
        const scopes = [{ items: [{ price: 10 }, { price: 20 }] }];
        expect(_renderInterpolatedString('Total: {{#formula SUM items.price}}', scopes))
            .toBe('Total: 30');
    });
});

// ============================================
// 7. renderCell
// ============================================

describe('renderCell', () => {
    it('should render simple value placeholder', () => {
        const scopes = [{ name: 'Test' }];
        expect(_renderCell('{{name}}', scopes)).toBe('Test');
    });

    it('should render interpolated string', () => {
        const scopes = [{ a: 'X', b: 'Y' }];
        expect(_renderCell('{{a}} and {{b}}', scopes)).toBe('X and Y');
    });

    it('should pass through non-placeholder values', () => {
        expect(_renderCell('hello', [{}])).toBe('hello');
        expect(_renderCell(42, [{}])).toBe(42);
    });

    it('should pass options to renderValue', () => {
        const scopes = [{ price: '100' }];
        expect(_renderCell('{{price}}', scopes, {})).toBe('100');
        expect(_renderCell('{{price}}', scopes, { autoParseNumbers: true })).toBe(100);
    });
});

// ============================================
// 8. calculateFormulaValue
// ============================================

describe('calculateFormulaValue', () => {
    it('should calculate SUM', () => {
        const scopes = [{ items: [{ price: 10 }, { price: 20 }] }];
        expect(_calculateFormulaValue('SUM', 'items', 'price', scopes)).toBe('30');
    });

    it('should calculate AVERAGE', () => {
        const scopes = [{ items: [{ price: 10 }, { price: 30 }] }];
        expect(_calculateFormulaValue('AVERAGE', 'items', 'price', scopes)).toBe('20');
    });

    it('should calculate COUNT', () => {
        const scopes = [{ items: [{ price: 10 }, { price: 20 }, { price: 30 }] }];
        expect(_calculateFormulaValue('COUNT', 'items', 'price', scopes)).toBe('3');
    });

    it('should return "0" for empty array', () => {
        const scopes = [{ items: [] }];
        expect(_calculateFormulaValue('SUM', 'items', 'price', scopes)).toBe('0');
    });
});

// ============================================
// 9. Integration: renderExcelTemplateFromBuffer
// ============================================

describe('renderExcelTemplateFromBuffer', () => {
    async function createTemplate(
        setup: (ws: ExcelJS.Worksheet) => void
    ): Promise<Buffer> {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        setup(ws);
        const arrayBuffer = await wb.xlsx.writeBuffer();
        return Buffer.from(arrayBuffer);
    }

    async function readOutput(buffer: Buffer): Promise<ExcelJS.Workbook> {
        const wb = new ExcelJS.Workbook();
        const ab = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
        await wb.xlsx.load(ab as ArrayBuffer);
        return wb;
    }

    it('should fill simple value placeholders', async () => {
        const template = await createTemplate((ws) => {
            ws.getCell('A1').value = '{{company}}';
            ws.getCell('B1').value = '{{invoice}}';
        });

        const result = await renderExcelTemplateFromBuffer(template, {
            company: 'Acme Corp',
            invoice: 'INV-001',
        });

        const wb = await readOutput(result);
        const ws = wb.getWorksheet('Sheet1')!;
        expect(ws.getCell('A1').value).toBe('Acme Corp');
        expect(ws.getCell('B1').value).toBe('INV-001');
    });

    it('should preserve numbers without auto-parse by default', async () => {
        const template = await createTemplate((ws) => {
            ws.getCell('A1').value = '{{code}}';
        });

        const result = await renderExcelTemplateFromBuffer(template, {
            code: '001234',
        });

        const wb = await readOutput(result);
        const ws = wb.getWorksheet('Sheet1')!;
        expect(ws.getCell('A1').value).toBe('001234');
    });

    it('should auto-parse numbers when option is enabled', async () => {
        const template = await createTemplate((ws) => {
            ws.getCell('A1').value = '{{price}}';
        });

        const result = await renderExcelTemplateFromBuffer(
            template,
            { price: '100' },
            { autoParseNumbers: true }
        );

        const wb = await readOutput(result);
        const ws = wb.getWorksheet('Sheet1')!;
        expect(ws.getCell('A1').value).toBe(100);
    });

    it('should render table with array data', async () => {
        const template = await createTemplate((ws) => {
            ws.getCell('A1').value = '{{#table items}}';
            ws.getCell('A2').value = '{{name}}';
            ws.getCell('B2').value = '{{qty}}';
        });

        const result = await renderExcelTemplateFromBuffer(template, {
            items: [
                { name: 'Widget A', qty: 5 },
                { name: 'Widget B', qty: 3 },
            ],
        });

        const wb = await readOutput(result);
        const ws = wb.getWorksheet('Sheet1')!;
        expect(ws.getCell('A1').value).toBe('Widget A');
        expect(ws.getCell('B1').value).toBe(5);
        expect(ws.getCell('A2').value).toBe('Widget B');
        expect(ws.getCell('B2').value).toBe(3);
    });

    it('should render interpolated strings', async () => {
        const template = await createTemplate((ws) => {
            ws.getCell('A1').value = 'Hello {{name}}, order {{id}}';
        });

        const result = await renderExcelTemplateFromBuffer(template, {
            name: 'John',
            id: 'ORD-001',
        });

        const wb = await readOutput(result);
        const ws = wb.getWorksheet('Sheet1')!;
        expect(ws.getCell('A1').value).toBe('Hello John, order ORD-001');
    });

    it('should handle join in cells', async () => {
        const template = await createTemplate((ws) => {
            ws.getCell('A1').value = '{{tags | join(", ")}}';
        });

        const result = await renderExcelTemplateFromBuffer(template, {
            tags: ['urgent', 'important', 'pending'],
        });

        const wb = await readOutput(result);
        const ws = wb.getWorksheet('Sheet1')!;
        expect(ws.getCell('A1').value).toBe('urgent, important, pending');
    });
});
