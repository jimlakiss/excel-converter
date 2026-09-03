# CostX Converter

A small static browser app for converting a CostX-exported spreadsheet into a cleaned import file for builder or project management software. It runs entirely in the browser using plain HTML, CSS, and JavaScript, with [SheetJS](https://sheetjs.com/) for flat imports and [ExcelJS](https://github.com/exceljs/exceljs) for template-preserving workbook exports.

## Features

- Static single-page interface with no backend, auth, or database
- Upload support for `.xlsx`, `.xls`, and `.csv`
- Flexible header mapping for common CostX-style columns
- Canonical row transformation pipeline for normalisation and validation
- Generic PM Import exporter implemented
- Wonderbuild exporter implemented from the portal column mapping
- Buildxact exporter implemented from the supplied import template
- Optional default markup percentage for Buildxact and Wonderbuild exports
- Optional description max length override with a default of 250 characters
- ANZSMM 2022 workbook exporter using the bundled standard template
- In-browser file processing and immediate download
- Results summary, warnings list, rejected rows list, and a 200-row preview

## Project Structure

- `index.html` - page structure and CDN script loading
- `style.css` - responsive layout and presentation styles
- `app.js` - parsing, transformation, validation, exporting, and UI logic
- `templates/ANZSMM 2022 - TEMPLATE_V1.0.xlsx` - base workbook for the ANZSMM exporter
- `README.md` - setup, deployment, and extension notes

## How It Works

1. The user uploads an Excel or CSV file.
2. If the workbook has multiple sheets, the app treats the first sheet as a cover/summary sheet and processes the remaining trade sheets.
3. If the workbook has only one sheet, it processes that sheet.
4. It detects the most likely header row in each processed sheet and maps known aliases to canonical fields.
5. Rows are normalised into a shared internal structure:
   - `code`
   - `description`
   - `category`
   - `subcategory`
   - `quantity`
   - `uom`
   - `rate`
   - `amount`
   - `notes`
   - `source_row_number`
   - `warnings`
6. The selected exporter converts canonical rows into the requested import sheet.
7. The browser immediately downloads the converted file.

### ANZSMM 2022 Workbook Export

The `ANZSMM 2022 Workbook` option starts from
`templates/ANZSMM 2022 - TEMPLATE_V1.0.xlsx`, then matches CostX trade tabs to
the template by trade name instead of tab number. This allows an export to use
only the standard trades present in the uploaded CostX workbook.

- All template tabs remain in the output workbook.
- Matching CostX tabs populate the corresponding template tab.
- Template-only tabs are blank below their fixed title and header rows.
- Exact final `Total` rows from CostX are omitted; other non-blank rows are copied.
- The template's summary sheet, formulas, formatting, and internal `HOME` links are retained.
- Populated trade tabs receive line-total and final-total formulas.
- The download uses the original input filename with `_iqs.xlsx` appended.

The template is loaded as a static site asset. Use GitHub Pages or a local web
server for this exporter rather than opening `index.html` directly with `file://`.

## Transformation Rules

- Trims extra whitespace
- Skips completely blank rows
- Skips subtotal, total, and section-header style rows where practical
- Normalises units:
  - `nr` -> `Item`
  - `no` -> `Item`
  - `each` -> `Item`
  - `m²` / `m2` -> `m2`
  - `m³` / `m3` -> `m3`
- Converts numeric strings to numbers where possible
- Preserves original worksheet row number
- Adds warnings for:
  - missing code
  - missing quantity
  - missing UoM
  - invalid quantity, rate, or amount values

## Run Locally

Because this is a static app, you can run it with any local static file server.

### Option 1: Open Directly

Open `index.html` in a browser.

Note: some browsers apply stricter local-file rules. If the spreadsheet library fails to load from the CDN, use a local server instead.

### Option 2: Use a Simple Local Server

If you have Python installed:

```bash
python3 -m http.server 8000
```

Then open:

```text
http://localhost:8000
```

## Deploy to GitHub Pages

1. Push the project to a GitHub repository.
2. In the repository, open `Settings`.
3. Go to `Pages`.
4. Under `Build and deployment`, choose `Deploy from a branch`.
5. Select your main branch and the root folder.
6. Save the settings.
7. GitHub Pages will publish the static site and serve `index.html`.

## Exporters

### Implemented

- `Generic PM Import`
- `Buildxact`
- `Wonderbuild`
- `ANZSMM 2022 Workbook`

Output columns:

- `Code`
- `Description`
- `Quantity`
- `UoM`
- `Rate`
- `Amount`
- `Category`
- `Notes`

Wonderbuild output columns:

- `Item Number`
- `Category`
- `Costing Item`
- `Location`
- `Quantity`
- `Wastage (%)`
- `Rounding`
- `UOM`
- `Cost (ex.)`
- `Cost Type`
- `Markup (%)`
- `Allowance Type (PC/PS)`
- `GST Free (Yes/No)`
- `Note`
- `SKU`

Buildxact output columns:

- `Category`
- `Description`
- `Quantity`
- `UOM`
- `UnitCost`
- `ItemType`
- `MarkupPercent`
- `ItemCode`

### Placeholders

- None yet beyond future custom exporters

## Extension Points

The code is intentionally split into small functions so future changes stay localised:

- File parsing: update `parseInputFile`
- Multi-sheet workbook selection: update `buildSheetContexts`
- Header alias matching: update `HEADER_ALIASES`
- Canonical row logic: update `createCanonicalRow`
- Validation rules: update `validateCanonicalRow`
- Export formats: add a new exporter and register it in `exportRows`
- Template exports: follow `exportAnzsmmWorkbook` and bundle each base workbook as a static asset
- UI rendering: extend `renderResults` and `renderPreviewTable`

## Future Extension Points

### Authentication

Authentication could be added later by placing the static frontend behind:

- a Rails app with session-based auth
- a small API gateway with token-based auth
- a hosted frontend using an identity provider such as Auth0 or Clerk

At that point, uploaded files could still be processed client-side, or the conversion could move server-side for auditing, job tracking, or more complex mapping workflows.

### Rails Integration

A Rails implementation could evolve this prototype into:

- a signed-in dashboard for estimators or admins
- stored mapping templates per client or software target
- background jobs for large workbook processing
- import history and downloadable audit logs
- customer-specific exporters for Buildxact, Wonderbuild, or other PM platforms

One clean next step would be to keep the current browser transformer as a reference implementation while extracting the mapping and exporter rules into shared JSON or Ruby service objects.

## Notes

- Flat import exports use the first worksheet as a summary tab when multiple sheets are present.
- The ANZSMM exporter maps all CostX trade worksheets to the standard template.
- No files are uploaded to a server.
- No analytics or third-party APIs are used.
- The app is designed for maintainability first, not visual complexity.
