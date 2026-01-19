# RETIREMENT PLANNING DOCUMENT GENERATOR
# Template-Based Multi-Source JSON System

## Overview

This system generates professional Word documents for retirement planning using:
1. **Template Structure** - Static document content with placeholders
2. **Multiple JSON Data Sources** - Dynamic data organized by source

## Architecture

```
data/
├── template_structure.json    # Document template with placeholders
├── cfr_data.json             # GREEN - Client Financial Record
├── cyc_data.json             # YELLOW - Calculations
├── illustration_data.json    # PINK - Illustration data
├── ceding_info.json          # RED - Ceding Info Check
└── user_input.json           # BLUE - User Input

generated_documents/          # Output folder for generated documents
```

## Data Sources (Color Mapping)

| Color  | Source          | JSON File              | Description                        |
|--------|-----------------|------------------------|------------------------------------|
| GREEN  | CFR             | cfr_data.json          | Client Financial Record data       |
| YELLOW | CYC             | cyc_data.json          | Calculation results                |
| PINK   | Illustration    | illustration_data.json | Illustration document data         |
| RED    | Ceding Info     | ceding_info.json       | Ceding information/file notes      |
| BLUE   | User Input      | user_input.json        | Adviser-entered client information |

## Placeholder Syntax

In the template, placeholders use the format:
```
{source_name.path.to.value}
```

Examples:
- `{cfr.recipient.title_and_name}` → Client name from CFR
- `{user_input.letter_details.date}` → Date from user input
- `{cyc.outperformance_table.scottish_widows.level_of_outperformance}` → Calculation result

## Files

| File                      | Description                                      |
|---------------------------|--------------------------------------------------|
| `generate_from_template.py` | **NEW** Main generator using multi-source JSON |
| `generate_refactored.py`  | Legacy generator (single JSON file)              |
| `create_template.py`      | Creates docx template with Jinja2 placeholders   |
| `style_helpers.py`        | Styling functions (fonts, margins, tables)       |

## Usage

### Generating a Document

1. **Edit the data files** in `data/` folder with your client data
2. **Run the generator**:
   ```cmd
   python generate_from_template.py
   ```
3. **Find output** in `generated_documents/template_output_YYYYMMDD_HHMMSS.docx`

### Modifying Template Content

Edit `data/template_structure.json` to change:
- Static text paragraphs
- Section headings
- Bullet lists
- Table structures

### Adding New Placeholders

1. Add the data field to the appropriate JSON file (e.g., `user_input.json`)
2. Reference it in the template using `{source.path.to.field}`
3. The generator will automatically resolve it

## Styling

Styles are defined in `template_structure.json`:

```json
"styles": {
  "document": { "font_family": "Poppins", "font_size": 10 },
  "heading1": { "font_family": "Noe Display SJP Bold", "font_size": 20 },
  "heading2": { "font_family": "Poppins SemiBold", "font_size": 12 },
  "table_header": { "background_color": "3FDCC8", "bold": true }
}
```

## Requirements

```
python-docx>=0.8.11
```

Install with:
```cmd
pip install -r requirements.txt
```

## Workflow

1. **Receive client data** from various sources
2. **Populate JSON files** with the data
3. **Run generator** → produces formatted .docx
4. **Review output** and adjust as needed

---

Last Updated: January 2026
