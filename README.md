# Plant Protection Evaluation Report Generator

This repository contains a small document automation prototype for generating a regulator-style DOCX report from structured JSON data. It simulates an ActiveDocs-style workflow with Python tools that are available locally and easy to inspect.

The primary workflow is a manually maintained Word template rendered with `docxtpl`. Python is responsible for validation, mapping, calculations, chart generation, and rendering. A code-generated template scaffold is still available as an optional fallback when the template needs to be bootstrapped from scratch.

## Technology Stack

- Python 3.12
- Pydantic for JSON schema validation and numeric constraints
- docxtpl and Jinja2 for DOCX template rendering
- python-docx for optional DOCX template scaffolding and Word-level checks
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

If the template already exists, the command uses it as-is. If the template is missing, a baseline version is scaffolded automatically so the pipeline still runs from a clean clone.

To regenerate the baseline template scaffold explicitly:

```powershell
python -m report_generator --rebuild-template
```

Use that option only when you want to recreate the code-generated starting point. If you have manually styled `templates/report_template.docx` in Word, run `python -m report_generator` without `--rebuild-template` so your manual template changes are preserved.

To use a different JSON input:

```powershell
python -m report_generator  --template templates\manual_template.docx --output generated\custom_report.docx
```

## CLI Reference

```
python -m report_generator [OPTIONS]
```

| Option | Default | Description |
|---|---|---|
| `--input PATH` | `data/input_data.json` | Path to the JSON input file. |
| `--template PATH` | `templates/report_template.docx` | Path to the DOCX template to render. |
| `--output PATH` | `generated/report.docx` | Destination path for the generated report. |
| `--charts-dir PATH` | `build/charts` | Directory where chart images are written. |
| `--rebuild-template` | *(flag)* | Regenerate the baseline template from code before rendering. Overwrites the existing template file. |

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
  -> manually authored Word template
  -> DOCX template rendering with docxtpl/Jinja2
  -> generated report.docx
```

The implementation keeps a clear boundary between data preparation and document layout. Python validates, normalizes, calculates, and maps data into display-ready values. The DOCX template handles layout, repeated product sections, tables, simple conditions, and chart placement.

## Template Workflow

The preferred approach for this project is:

- maintain `templates/report_template.docx` manually in Word for layout and styling
- keep logic, flags, formatting, and calculations in Python
- use `docxtpl` only to render placeholders, simple loops, and simple conditions

The code-generated template scaffold exists for convenience:

- it provides a runnable baseline from a clean clone
- it captures a fallback layout in source control
- it should not be treated as the primary template-authoring workflow once the Word template is being maintained manually

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

The maintained `templates/report_template.docx` includes:

- Title page
- Header with authority and report ID
- Footer with classification, generator version, disclaimer, and Word page fields
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

## Generator Versioning

The package version is declared in `report_generator/__init__.py`:

```python
__version__ = "1.0.0"
```

This value is injected into every generated report via the footer (`INTERNAL · v1.0.0`). Bumping `__version__` before a run is enough to make all subsequent reports traceable to the exact code that produced them. No other files need to change.

## Known Limitations

- Automated tests are intentionally skipped for now, per the current implementation request.
- The template can be scaffolded programmatically for reproducibility, but the intended primary workflow is manual template authoring in Word.
- Temporary chart image files are ignored by Git because they are regenerated on every run and embedded in the DOCX output.
