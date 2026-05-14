# Plant Protection Evaluation Report Generator

This repository contains a small document automation prototype for generating a regulator-style DOCX report from structured JSON data. It simulates an ActiveDocs-style workflow with Python tools that are available locally and easy to inspect.

## Technology Stack

- Python 3.13
- Pydantic for JSON schema validation and numeric constraints
- docxtpl and Jinja2 for DOCX template rendering
- python-docx for generating the reusable DOCX template structure
- matplotlib for efficacy charts

## One-Command Run

From the repository root:

```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m report_generator
```

The command creates or updates these outputs:

- `templates/report_template.docx`
- `generated/report.docx`
- `build/charts/*.png`

To rebuild the template explicitly:

```powershell
python -m report_generator --rebuild-template
```

To use a different JSON input:

```powershell
python -m report_generator --input path\to\input_data.json --output generated\custom_report.docx
```

## Input and Output Paths

- Input data: `data/input_data.json`
- DOCX template: `templates/report_template.docx`
- Generated report: `generated/report.docx`
- Temporary chart images: `build/charts/`

## Architecture

```text
JSON input
  -> load_json()
  -> Pydantic schema validation
  -> Python business rules and report view-model mapping
  -> matplotlib chart generation
  -> DOCX template rendering with docxtpl/Jinja2
  -> generated report.docx
```

The implementation keeps a clear boundary between data preparation and document layout. Python validates, normalizes, calculates, and maps data into display-ready values. The DOCX template handles layout, repeated product sections, tables, simple conditions, and chart placement.

## Business Rules

- A high-risk notice is shown when `risk_score > 70`.
- Restrictions are shown only when `approval_status == "conditional"`.
- German wording is used when `country == "DE"`.
- Product average efficacy is the mean of its application efficacy values.
- Overall average efficacy is the mean of product-level averages.
- Products with no applications remain in the report and show fallback text instead of a chart.

For the supplied data this means:

- AgriShield 250 EC shows the high-risk notice, restrictions, and German approval wording.
- Verdura 75 WG does not show the high-risk notice or restrictions.
- Overall average efficacy is calculated as 86.3%.

## Validation and Defensive Behavior

The generator validates:

- Required structural fields such as report metadata, product name, approval period, risk components, and toxicity indicators.
- Risk scores and component scores in the range 0-100.
- Efficacy percentages in the range 0-100.
- Positive dose and toxicity indicator values.
- Non-negative maximum applications and preharvest interval days.
- Controlled approval statuses: `approved`, `conditional`, `rejected`, `expired`.

Rendering uses Jinja2 `StrictUndefined`, so missing template variables fail fast. After rendering, the generated DOCX is checked for unresolved Jinja markers such as `{{` or `{%`.

## Template Design

The generated `templates/report_template.docx` includes:

- Title page
- Header with authority and report ID
- Footer with classification, disclaimer, and Word page fields
- Executive summary table
- Reusable product section rendered once per product
- Applications table
- Risk components table
- Toxicity indicators table
- Efficacy chart placeholder
- Conditional high-risk notice
- Conditional restrictions section
- Environmental notes and references

The main reusable component is the product loop. This mirrors how the same design would become a reusable subtemplate or component in an enterprise document automation tool.

## ActiveDocs Translation

| Prototype concept | ActiveDocs equivalent |
|---|---|
| `input_data.json` | External data source / JSON / XML input |
| Pydantic models | Input schema and data contract |
| Python mapper | Data preparation and transformation layer |
| Jinja conditions | ActiveDocs business rules |
| Jinja loops | Repeating sections |
| DOCX template | ActiveDocs template |
| Product loop | Reusable subtemplate/component |
| Generated DOCX | Generated document package |

The same principle applies in both approaches: validate and reshape data before rendering, keep business rules explicit and testable, and keep template sections reusable.

## Known Limitations

- PDF export is not included because reliable conversion usually depends on Word, LibreOffice, or a dedicated document service being available in the runtime environment.
- Automated tests are intentionally skipped for now, per the current implementation request.
- The template is generated programmatically so the project is reproducible without manual Word editing. It can still be opened and styled further in Word.
- Temporary chart image files are ignored by Git because they are regenerated on every run and embedded in the DOCX output.
