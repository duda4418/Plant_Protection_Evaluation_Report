# Assignment Implementation Plan — Dynamic Regulatory Report Generator

## 1. Goal

Build a small document automation prototype that reads `input_data.json` and generates a regulator-style report in DOCX format.

The solution should simulate an ActiveDocs-style workflow without using ActiveDocs:

```text
JSON input
  -> validation
  -> mapping / business rules
  -> chart generation
  -> DOCX template rendering
  -> generated report
```

The assignment requires a generated report, template files, source code/configuration, README, and 5–10 short slides explaining the approach, design, business logic, scaling, and how the solution would translate to ActiveDocs.

---

## 2. Required Features

Implement the following:

- Read `input_data.json`.
- Generate a regulator-style report.
- Include a header and footer on every page.
- Show a high-risk notice when `risk_score > 70`.
- Show a restrictions section when `approval_status == "conditional"`.
- Use German wording when `country == "DE"`.
- Render tables for:
  - applications
  - toxicity indicators
  - risk components
- Generate at least one plot, preferably efficacy by crop.
- Use at least one reusable template component.
- Provide a one-command run from the README.
- Commit the generated output to the repository.

---

## 3. Recommended Technology Stack

Use the following Python stack:

```text
Python
docxtpl
Jinja2
Pydantic
matplotlib
python-docx(only if necesarry)
pytest
```

### Why these tools?

#### `docxtpl`

Used to render a Microsoft Word `.docx` template with dynamic values.

Example inside the Word template:

```jinja
Product: {{ product.name }}

{% if product.show_high_risk_notice %}
HIGH RISK NOTICE
{% endif %}
```

#### `Jinja2`

The templating language used inside the DOCX template.

Use it for:

- variables: `{{ product.name }}`
- simple conditions: `{% if product.show_restrictions %}`
- loops: `{% for app in product.applications %}`

#### `Pydantic`

Used to validate the JSON input before generating the report.

This protects the generator from invalid data such as:

- missing required fields
- invalid risk scores
- invalid efficacy percentages
- wrong data types
- negative values

#### `matplotlib`

Used to generate efficacy charts as image files.

#### `python-docx`

Used mainly for testing generated DOCX content and optionally for advanced document checks.

#### `pytest`

Used for automated tests.

---

## 4. Proposed Project Structure

```text
regulatory-report-generator/
│
├── data/
│   └── input_data.json
│
├── templates/
│   └── report_template.docx
│
├── generated/
│   ├── report.docx
│   └── report.pdf
│
├── build/
│   └── charts/
│
├── src/
│   └── report_generator/
│       ├── __init__.py
│       ├── cli.py
│       ├── models.py
│       ├── mapper.py
│       ├── charts.py
│       ├── renderer.py
│       └── formatting.py
│
├── tests/
│   ├── test_validation.py
│   ├── test_mapping.py
│   ├── test_business_rules.py
│   └── test_rendering.py
│
├── requirements.txt
├── README.md
├── ASSIGNMENT.md
└── slides/
    └── approach.pdf
```

---

## 5. Architecture

### 5.1 Data Flow

```text
input_data.json
  -> load_json()
  -> Pydantic schema validation
  -> Python domain models
  -> report view-model mapping
  -> chart generation
  -> DOCX template rendering
  -> report.docx
  -> optional report.pdf
```

### 5.2 Separation of Concerns

Keep a clear boundary between code and template.

#### Python should handle:

- JSON loading
- schema validation
- data normalization
- calculations
- business rules
- chart generation
- fallback handling
- error messages

#### DOCX template should handle:

- layout
- static text
- placeholders
- simple loops
- simple conditions
- tables
- chart placement

Do not put complex business logic inside the DOCX template.

---

## 6. Implementation Steps

### Step 1 — Set up the project

Create the folder structure and install dependencies.

```bash
pip install docxtpl pydantic matplotlib python-docx pytest
```

Create a runnable command, for example:

```bash
python -m report_generator --input data/input_data.json --template templates/report_template.docx --output generated/report.docx
```

Optional:

```bash
make report
```

---

### Step 2 — Implement JSON loading

Create a safe JSON loader.

Requirements:

- check that the input file exists
- use UTF-8
- catch invalid JSON
- show clear errors

Example behavior:

```text
Invalid JSON in data/input_data.json: expected ',' at line 12
```

---

### Step 3 — Implement Pydantic validation

Create models for:

- report metadata
- product
- concentration
- approval period
- risk components
- application
- toxicity indicators
- references
- footer

Validation rules:

```text
risk_score: 0-100
efficacy_percent: 0-100
dose: > 0
max_applications_per_season: >= 0
preharvest_interval_days: >= 0
country: non-empty string
approval_status: controlled value such as approved/conditional/rejected/expired
```

Fail fast when required data is invalid.

---

### Step 4 — Implement business rules

Create explicit, testable functions.

```python
def should_show_high_risk_notice(risk_score: int) -> bool:
    return risk_score > 70


def should_show_restrictions(approval_status: str) -> bool:
    return approval_status.lower() == "conditional"


def is_german_country(country: str) -> bool:
    return country.upper() == "DE"
```

Expected behavior from the provided data:

```text
AgriShield 250 EC
- risk_score = 78
- approval_status = conditional
- country = DE
- should show high-risk notice
- should show restrictions
- should use German wording

Verdura 75 WG
- risk_score = 42
- approval_status = approved
- country = FR
- should not show high-risk notice
- should not show restrictions
- should use normal country/status wording
```

---

### Step 5 — Implement calculations

Calculate product-level average efficacy.

```python
def calculate_product_average_efficacy(applications: list[Application]) -> float | None:
    if not applications:
        return None

    values = [app.efficacy_percent for app in applications]
    return sum(values) / len(values)
```

Calculate overall average efficacy as the mean of product-level averages.

This matches the supplied example output:

```text
AgriShield average: 83.7%
Verdura average: 89.0%
Overall average: 86.3%
```

Document this assumption in the README.

---

### Step 6 — Implement formatting helpers

Create formatting functions instead of formatting directly inside the template.

Examples:

```python
def format_percent(value: float | None, decimals: int = 1) -> str:
    if value is None:
        return "N/A"
    return f"{value:.{decimals}f}%"


def format_dose(value: float, unit: str) -> str:
    return f"{value:g} {unit}"


def format_concentration(value: float, unit: str) -> str:
    return f"{value:g} {unit}"


def format_classification(value: str) -> str:
    return value.upper()
```

---

### Step 7 — Implement the mapping layer

Do not pass raw JSON directly into the DOCX template.

Create a clean report context.

Example:

```python
context = {
    "metadata": {
        "report_title": "...",
        "report_id": "...",
        "issuing_authority": "...",
        "issued_on": "...",
        "version": "...",
        "classification": "INTERNAL",
    },
    "summary": {
        "product_count": 2,
        "conditional_count": 1,
        "overall_average_efficacy": "86.3%",
    },
    "products": [
        {
            "name": "AgriShield 250 EC",
            "risk_score": 78,
            "show_high_risk_notice": True,
            "show_restrictions": True,
            "approval_line": "Germany: Zugelassen gem. PflSchG — Status: conditional. Valid from 2025-01-01 to 2027-12-31.",
            "application_count": 3,
            "average_efficacy": "83.7%",
            "applications": [...],
            "toxicity_rows": [...],
            "risk_component_rows": [...],
            "restrictions": [...],
            "environmental_notes": [...],
        }
    ],
}
```

The template should receive display-ready values.

---

### Step 8 — Generate charts

Generate one efficacy chart per product.

Recommended chart:

```text
Horizontal bar chart
Y-axis: crop
X-axis: efficacy percent
Range: 0-100
```

Defensive behavior:

- if a product has no applications, skip the chart
- display `No efficacy data available` in the report
- do not crash

Store chart images in:

```text
build/charts/
```

---

### Step 9 — Create the DOCX template manually

Create `templates/report_template.docx` manually in Word or LibreOffice.

The template should contain:

- title page
- header with authority and report ID
- footer with classification and page numbers
- overall summary
- reusable product section
- applications table
- toxicity indicators table
- risk components table
- efficacy chart placeholder
- conditional high-risk notice
- conditional restrictions section
- environmental notes

---

## 7. DOCX Template Best Practices

### 7.1 Keep placeholders simple

Good:

```jinja
{{ product.name }}
{{ product.average_efficacy }}
{{ metadata.report_id }}
```

Avoid:

```jinja
{{ products[0]["risk_components"]["human_health"] }}
```

---

### 7.2 Keep business logic out of the template

Bad:

```jinja
{% if product.country == "DE" and product.approval_status == "conditional" and product.risk_score > 70 %}
```

Good:

```jinja
{% if product.show_high_risk_notice %}
```

Compute complex conditions in Python.

---

### 7.3 Do not split placeholders across Word formatting

Avoid bolding only part of a placeholder.

Bad:

```text
{{ product.
name }}
```

Bad:

```text
{{ product.name }}
```

where only part of the placeholder is bold.

Good:

```text
{{ product.name }}
```

typed as one continuous placeholder.

---

### 7.4 Use table row loops carefully

Use Jinja row loops for dynamic tables.

Example concept:

```jinja
{% for app in product.applications %}
{{ app.crop }} | {{ app.target_pest }} | {{ app.bbch }} | {{ app.dose }} | {{ app.efficacy }} | {{ app.phi_days }}
{% endfor %}
```

For real Word tables, place loop tags in table rows according to `docxtpl` table syntax.

---

### 7.5 Use dynamic page fields

Use Word fields for page numbering:

```text
Page { PAGE } / { NUMPAGES }
```

Do not hardcode page numbers.

---

### 7.6 Use reusable sections

The product section should be reusable:

```jinja
{% for product in products %}
    product section
{% endfor %}
```

This is the main reusable template component.

---

## 8. Defensive Programming Guidelines

### 8.1 Never trust the input JSON

Validate everything before rendering.

Fail clearly if required data is invalid.

---

### 8.2 Normalize casing

Handle:

```text
DE
de
De

conditional
Conditional
CONDITIONAL
```

Normalize in Python:

```python
country = country.upper()
approval_status = approval_status.lower()
```

---

### 8.3 Handle empty lists

The generator should not crash when these are empty:

```text
products
applications
restrictions
environmental_notes
references
```

Decide behavior:

```text
products empty -> fail with clear error or generate "No products available"
applications empty -> no chart, show "No applications provided"
restrictions empty -> omit section or show "No restrictions provided"
environmental notes empty -> omit section or show "No environmental notes provided"
```

Document your choice.

---

### 8.4 Validate numeric ranges

Reject invalid values such as:

```text
risk_score = -1
risk_score = 120
efficacy_percent = 140
dose = 0
preharvest_interval_days = -5
```

---

### 8.5 Use StrictUndefined

Configure Jinja to fail if the template references a missing field.

```python
from jinja2 import Environment, StrictUndefined

jinja_env = Environment(undefined=StrictUndefined)
doc.render(context, jinja_env=jinja_env)
```

This prevents silent empty values in the report.

---

### 8.6 Check for unresolved placeholders

After rendering, test that the generated DOCX does not contain:

```text
{{
}}
{%
%}
```

---

### 8.7 Avoid hardcoding sample products

Do not write code that assumes exactly two products.

Bad:

```python
product_1 = data["products"][0]
product_2 = data["products"][1]
```

Good:

```python
for product in report.products:
    ...
```

---

## 9. Edge Cases to Test

Create tests or sample JSON variants for:

```text
1. Empty products list
2. Product with no applications
3. Product with risk_score exactly 70
4. Product with risk_score 71
5. Product with country "de" lowercase
6. Conditional product with empty restrictions
7. Approved product with restrictions
8. Very long crop and pest names
9. Invalid efficacy_percent > 100
10. Missing required product_name
11. Invalid JSON file
12. Unknown country code
13. Missing environmental notes
14. Large number of products
15. Large number of applications
```

Expected behavior:

```text
Either generate a valid report or fail with a clear validation error.
Never generate a corrupted DOCX.
Never crash with an unclear stack trace.
```

---

## 10. Testing Strategy

### 10.1 Unit tests

Test pure Python logic:

- validation
- business rules
- averages
- formatting
- German wording
- restrictions logic

Example:

```python
def test_high_risk_notice_threshold():
    assert should_show_high_risk_notice(71) is True
    assert should_show_high_risk_notice(70) is False
```

---

### 10.2 Mapping tests

Verify raw input becomes the correct report context.

```python
def test_agri_shield_mapping():
    product = map_product(...)
    assert product["show_high_risk_notice"] is True
    assert product["show_restrictions"] is True
    assert "Zugelassen gem. PflSchG" in product["approval_line"]
```

---

### 10.3 DOCX integration tests

Generate a real DOCX in a temporary directory and inspect the text.

```python
from docx import Document

def extract_docx_text(path):
    document = Document(path)
    paragraphs = [p.text for p in document.paragraphs]
    table_cells = [
        cell.text
        for table in document.tables
        for row in table.rows
        for cell in row.cells
    ]
    return "\n".join(paragraphs + table_cells)
```

Check:

```python
assert "Plant Protection Product Evaluation Report" in text
assert "AgriShield 250 EC" in text
assert "Verdura 75 WG" in text
assert "HIGH RISK NOTICE" in text
assert "Zugelassen gem. PflSchG" in text
assert "{{" not in text
assert "{%" not in text
```

---

## 11. README Requirements

The README should include:

```text
- project overview
- technology stack
- how to run
- how to run tests
- input/output paths
- architecture summary
- business rules
- assumptions
- defensive behavior
- known limitations
- how this would map to ActiveDocs
```

---

## 12. Assumptions to Document

Document assumptions clearly.

Recommended assumptions:

```text
- Overall average efficacy is calculated as the mean of product-level averages, matching the supplied example output.
- Restrictions are displayed only for conditional approvals.
- If a conditional product has no restrictions, the report displays a fallback message.
- Unknown country codes use a generic country/status wording.
- Products without applications are included, but their efficacy chart is omitted.
- Required structural fields fail validation if missing.
- Optional display sections may be omitted or replaced with "N/A".
```

---

## 13. ActiveDocs Translation

Explain how the prototype would translate to ActiveDocs.

```text
Prototype concept        ActiveDocs equivalent
------------------------------------------------
input_data.json          external data source / XML / JSON data
Pydantic models          input schema / data contract
Python mapper            data preparation layer
Jinja conditions         ActiveDocs business rules
Jinja loops              repeating sections
DOCX template            ActiveDocs template
Product section loop     reusable subtemplate/component
Generated DOCX/PDF       generated document package
```

Important message:

```text
The same architectural principle applies: reshape and validate data before rendering, keep reusable template sections, and keep business rules explicit and testable.
```

---

## 14. Known Risks and Mitigations

| Risk | Mitigation |
|---|---|
| Invalid JSON | Catch parse errors and print clear message |
| Missing required field | Pydantic validation |
| Invalid numeric values | Pydantic constraints |
| Empty applications | Skip chart and show fallback text |
| Missing template variable | Jinja StrictUndefined |
| Broken DOCX placeholders | Keep tags simple and in one Word run |
| Large data volume | Summarize in DOCX, export raw detail separately if needed |
| Page numbers not updating | Use Word fields and optional PDF export |
| Long text breaks tables | Enable wrapping and test long values |
| Hardcoded sample data | Loop over products generically |

---

## 15. Final Submission Checklist

Before submitting, verify:

```text
[ ] One-command run works from a clean clone
[ ] Generated DOCX is committed
[ ] Optional generated PDF is committed
[ ] Template DOCX is committed
[ ] README is complete
[ ] ASSIGNMENT.md is included
[ ] Tests pass
[ ] Business rules are covered by tests
[ ] Edge cases are documented
[ ] No unresolved placeholders in generated report
[ ] No hardcoded product-specific logic
[ ] Slides are included
```

---

## 16. Interview Talking Points

Use this explanation:

```text
I treated the assignment as a small document automation system, not just a script.
The JSON is validated first, then transformed into a report-specific view model.
Business rules are implemented in Python so they are testable.
The DOCX template remains focused on layout, placeholders, loops, and simple conditions.
The product section is reusable and rendered once per product.
The solution is defensive against missing data, empty collections, invalid numeric values, and template/data mismatches.
```