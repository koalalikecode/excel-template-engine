# 📊 Excel Template Engine

A lightweight library for rendering Excel templates with dynamic data using mustache-like `{{placeholder}}` syntax.

[![npm version](https://img.shields.io/npm/v/excel-template-engine.svg)](https://www.npmjs.com/package/excel-template-engine)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## ✨ Features

- 🔄 **Value Placeholders** - Fill cells with `{{path.to.value}}`
- 📋 **Table Rendering** - Expand arrays into multiple rows with `{{#table items}}`
- 🔗 **Join Arrays** - Combine arrays into strings with `{{tags | join(", ")}}`
- 📐 **Dynamic Formulas** - Generate Excel formulas like `=SUM(C5:C15)`
- 🎨 **Rich Text Support** - Preserve text formatting and styles
- 📄 **Auto Print Area** - Automatically adjusts print area when rows are added
- ✨ **Style Preservation** - Keeps cell styles, borders, and merge cells

## 📦 Installation

```bash
npm install excel-template-engine
```

## 🚀 Quick Start

### ESM (ES Modules)

```typescript
import { renderExcelTemplate } from 'excel-template-engine';

const data = {
  company: { name: 'Acme Corp' },
  invoice: { number: 'INV-001', date: '2024-01-15' },
  items: [
    { name: 'Widget A', quantity: 5, price: 100 },
    { name: 'Widget B', quantity: 3, price: 200 },
  ],
};

await renderExcelTemplate(
  './template.xlsx',
  data,
  './output.xlsx'
);
```

### CommonJS

```javascript
const { renderExcelTemplate } = require('excel-template-engine');

// Same usage as above
```

## 📖 Template Syntax

### 1. Value Placeholder

Fill a cell with a single value:

```
{{company.name}}          → "Acme Corp"
{{invoice.number}}        → "INV-001"
{{customer.address.city}} → "New York" (nested paths supported)
```

**Type preservation:**
- Numbers remain as Excel numbers
- Dates are converted to Excel dates
- ISO date strings (`2024-01-15`) are auto-parsed to dates

### 2. String Interpolation

Combine text with multiple placeholders in a single cell:

```
Hello {{customer.name}}, your order {{order.id}} is ready!
```

### 3. Table Rendering

Expand an array into multiple rows. Requires **2 consecutive rows**:

- **Row 1 (Anchor)**: Contains `{{#table arrayPath}}`
- **Row 2 (Template)**: Contains field placeholders like `{{fieldName}}`

**Template:**

| Row | A | B | C |
|-----|---|---|---|
| 5 | `{{#table items}}` | | |
| 6 | `{{name}}` | `{{quantity}}` | `{{price}}` |

**Data:**

```javascript
{
  items: [
    { name: 'Widget A', quantity: 5, price: 100 },
    { name: 'Widget B', quantity: 3, price: 200 },
    { name: 'Widget C', quantity: 2, price: 150 },
  ]
}
```

**Output:**

| Row | A | B | C |
|-----|---|---|---|
| 5 | Widget A | 5 | 100 |
| 6 | Widget B | 3 | 200 |
| 7 | Widget C | 2 | 150 |

> **Note:** The anchor row (`{{#table ...}}`) and template row are hidden after rendering.

### 4. Join Arrays

Combine array elements into a single cell:

```
{{tags | join(", ")}}           → "urgent, important, pending"
{{items.name | join(" / ")}}    → "Apple / Banana / Cherry"
```

**Supports:**
- Primitive arrays: `["a", "b", "c"]`
- Object arrays with property access: `items.name` extracts `name` from each item

### 5. Dynamic Formulas

Generate Excel formulas that reference table data. Supports 3 syntax styles:

#### Option 1: Column Letter (simplest)

Reference column directly:

| Syntax | Output |
|--------|--------|
| `{{#formula SUM C}}` | `=SUM(C5:C7)` |
| `{{#formula AVERAGE C}}` | `=AVERAGE(C5:C7)` |
| `{{#formula COUNT C}}` | `=COUNT(C5:C7)` |
| `{{#formula MIN C}}` | `=MIN(C5:C7)` |
| `{{#formula MAX C}}` | `=MAX(C5:C7)` |

> Uses the **last rendered table** in the worksheet.

#### Option 2: Field Name Only

Reference by field name (auto-detects column from template row):

```
{{#formula SUM price}}     → =SUM(C5:C7)  (if {{price}} is in column C)
{{#formula AVERAGE qty}}   → =AVERAGE(B5:B7)
```

> Uses the **last rendered table** in the worksheet.

#### Option 3: Array.Field Path (for multiple tables)

Explicitly specify which table to reference:

```
{{#formula SUM items.price}}      → =SUM(C5:C7)  (targets "items" table)
{{#formula SUM services.cost}}    → =SUM(D10:D15) (targets "services" table)
```

> Useful when you have **multiple tables** in the same worksheet.

#### Fallback Calculation

If no matching table is found, the formula calculates the value directly from data:

```
{{#formula SUM orders.total}}  → 450 (number, not formula)
```

#### Complete Example

**Template:**

| Row | A | B | C |
|-----|---|---|---|
| 5 | `{{#table items}}` | | |
| 6 | `{{name}}` | `{{qty}}` | `{{price}}` |
| 8 | **Total** | | `{{#formula SUM C}}` |

Or equivalently:
- `{{#formula SUM price}}`
- `{{#formula SUM items.price}}`

**Output:**

| Row | A | B | C |
|-----|---|---|---|
| 5 | Widget A | 5 | 100 |
| 6 | Widget B | 3 | 200 |
| 7 | Widget C | 2 | 150 |
| 8 | **Total** | | `=SUM(C5:C7)` |


## 🔧 API Reference

### `renderExcelTemplate(templatePath, data, outputPath)`

Render a template file and save to disk.

**Parameters:**

| Name | Type | Description |
|------|------|-------------|
| `templatePath` | `string` | Path to the `.xlsx` template file |
| `data` | `object` | Data object to fill the template |
| `outputPath` | `string` | Path for the output `.xlsx` file |

**Returns:** `Promise<void>`

```typescript
await renderExcelTemplate(
  './template.xlsx',
  data,
  './output.xlsx'
);
```

### `renderExcelTemplateFromBuffer(templateBuffer, data)`

Render from a buffer (useful for server-side applications).

**Parameters:**

| Name | Type | Description |
|------|------|-------------|
| `templateBuffer` | `Buffer \| ArrayBuffer \| Uint8Array` | Template file as buffer |
| `data` | `object` | Data object to fill the template |

**Returns:** `Promise<Buffer>`

```typescript
import { renderExcelTemplateFromBuffer } from 'excel-template-engine';
import fs from 'fs/promises';

const templateBuffer = await fs.readFile('./template.xlsx');
const outputBuffer = await renderExcelTemplateFromBuffer(templateBuffer, data);

// Send to client or save
await fs.writeFile('./output.xlsx', outputBuffer);
```

## 📝 TypeScript Support

Full TypeScript support with exported types:

```typescript
import { renderExcelTemplate } from 'excel-template-engine';
import type { Placeholder, FormulaType } from 'excel-template-engine';

// FormulaType: 'SUM' | 'AVERAGE' | 'COUNT' | 'MIN' | 'MAX'
```

## 📋 Full Example

**Template Structure:**

| Row | A | B | C | D |
|-----|---|---|---|---|
| 1 | **INVOICE** | | | |
| 2 | Company: | `{{company.name}}` | | |
| 3 | Invoice #: | `{{invoice.number}}` | Date: | `{{invoice.date}}` |
| 5 | `{{#table items}}` | | | |
| 6 | `{{name}}` | `{{description}}` | `{{quantity}}` | `{{price}}` |
| 8 | | | **Total:** | `{{#formula SUM D}}` |
| 10 | Notes: `{{notes | join("; ")}}` | | | |

**Data:**

```typescript
const data = {
  company: { name: 'Acme Corp' },
  invoice: { number: 'INV-2024-001', date: '2024-01-15' },
  items: [
    { name: 'Widget A', description: 'Premium widget', quantity: 5, price: 100 },
    { name: 'Widget B', description: 'Standard widget', quantity: 3, price: 75 },
    { name: 'Service', description: 'Installation', quantity: 1, price: 200 },
  ],
  notes: ['Handle with care', 'Express shipping', 'Gift wrap'],
};
```

**Output:**

A complete Excel file with:
- ✅ Company and invoice info filled in
- ✅ 3 item rows dynamically created
- ✅ Formula `=SUM(D5:D7)` auto-generated
- ✅ Notes joined as "Handle with care; Express shipping; Gift wrap"

## ⚠️ Limitations

| Limitation | Description |
|------------|-------------|
| `.xlsx` only | Does not support `.xls` (Excel 97-2003) |
| Formula placement | Formulas referencing a table must be placed **below** that table in the template |
| Single template row | Each table supports one template row (for multiple row patterns, use multiple tables) |

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

MIT
