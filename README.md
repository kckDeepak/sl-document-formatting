# SL Document Formatting

A template-based document generation system that creates professional Word documents for retirement planning. The system uses multiple JSON data sources to populate a templated document structure.

## Quick Start

### Clone and Setup

```bash
# Clone the repository
git clone https://github.com/kckDeepak/sl-document-formatting.git
cd sl-document-formatting

# Create virtual environment
python -m venv env

# Activate virtual environment
# Windows:
.\env\Scripts\Activate.ps1
# Linux/Mac:
source env/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Generate a Document

```bash
python generate_from_template.py
```

Output will be saved in `generated_documents/` folder.

---

## Architecture

```
sl-document-formatting/
├── generate_from_template.py    # Main generator script
├── style_helpers.py             # Styling utilities (fonts, tables, etc.)
├── requirements.txt             # Python dependencies
├── README.md
├── .gitignore
│
├── data/                        # JSON data sources
│   ├── template_structure.json  # Document template with placeholders
│   ├── cfr_data.json           # Client Financial Record (GREEN)
│   ├── cyc_data.json           # Calculations (YELLOW)
│   ├── illustration_data.json  # Illustration data (PINK)
│   ├── ceding_info.json        # Ceding Info/File notes (RED)
│   └── user_input.json         # User Input (BLUE)
│
├── images/
│   └── cover_page.png          # Full-page cover image
│
└── generated_documents/         # Output folder for generated .docx files
```

---

## Data Sources

The system pulls data from 5 JSON files, each representing a different source:

| Color  | File                     | Description                          |
|--------|--------------------------|--------------------------------------|
| 🟢 GREEN  | `cfr_data.json`         | Client Financial Record              |
| 🟡 YELLOW | `cyc_data.json`         | Calculation results                  |
| 🩷 PINK   | `illustration_data.json`| Illustration document data           |
| 🔴 RED    | `ceding_info.json`      | Ceding information / file notes      |
| 🔵 BLUE   | `user_input.json`       | Adviser-entered client information   |

---

## Placeholder Syntax

In the template, dynamic content uses placeholders in the format:

```
{source_name.path.to.value}
```

**Examples:**
- `{cfr.recipient.title_and_name}` → Client name from CFR
- `{user_input.letter_details.date}` → Date from user input
- `{cyc.outperformance_table.scottish_widows.level_of_outperformance}` → Calculation

---

## Document Structure

The generated document includes:

1. **Cover Page** - Full-page image
2. **Letter** - Introduction, objectives, documentation list
3. **Part 1** - Objectives, Needs and Circumstances
4. **Part 2** - Recommendation
5. **Part 3** - Impact of Replacement
6. **Part 4** - Attitude to Risk and Fund Selection
7. **Appendix i** - Further Details
8. **Appendix ii** - Product Comparison (landscape table)

---

## Customization

### Change Cover Image
Replace `images/cover_page.png` with your own image (A4 ratio recommended: 210mm × 297mm).

### Modify Template Text
Edit `data/template_structure.json` to change static content.

### Update Client Data
Edit the relevant JSON file in `data/` folder:
- Client details → `cfr_data.json`
- Calculations → `cyc_data.json` 
- User input → `user_input.json`

---

## Requirements

- Python 3.8+
- python-docx
- docxtpl
- Jinja2

See `requirements.txt` for specific versions.

---

## License

Private and Confidential
