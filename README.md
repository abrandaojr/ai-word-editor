# AI Word Editor

Python tool to revise a Microsoft Word document for a mixed audience: informed
non-specialists, policy readers, journalists, practitioners, and specialists who
expect technical accuracy.

The tool reads one `.docx` file and produces:

- Revised **Brazilian Portuguese** Word file and PDF
- Revised **American English** Word file and PDF
- LaTeX source files and extracted figures/tables used to build the PDFs
- Optional **visual communication audit** (DOCX + PDF) for all embedded figures, tables, and maps

## What it does

The editorial prompt applies principles from the following references (silently, without
mentioning them in the output):

- William Zinsser, *On Writing Well*
- Joseph M. Williams and Joseph Bizup, *Style: Lessons in Clarity and Grace*
- Helen Sword, *Stylish Academic Writing*
- Gerald Graff and Cathy Birkenstein, *They Say / I Say*
- MEAL Plan: Main idea, Evidence, Analysis, Link
- Randy Olson, *Houston, We Have a Narrative*
- Nancy Baron, *Escape from the Ivory Tower*
- Roy Peter Clark, *Writing Tools*
- *The Economist* style principles
- *The New York Times* explanatory journalism style
- Chip Heath and Dan Heath, *Made to Stick*
- Edward Tufte, Alberto Cairo, Cole Nussbaumer Knaflic, and Jonathan Schwabish for numerical and data communication

The visual audit prompt applies principles from:

- Tufte, Cairo, Knaflic, Schwabish (data visualization)
- Cynthia Brewer / ColorBrewer (cartographic palettes)
- Stephen Few (table and chart design)
- Nature and PNAS author figure guidelines
- Reuters Graphics and NYT Graphics standards

## Requirements

- Python 3.11+
- Anthropic API key

Optional for LaTeX/PDF output:

- `pdflatex` (MiKTeX or TeX Live). If unavailable, a ReportLab fallback PDF is generated automatically.

## Installation

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env
notepad .env
```

Add your Anthropic API key to `.env`:

```
ANTHROPIC_API_KEY=sk-ant-api03-your-real-key-here
```

## Usage

### API mode (default)

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Document.docx"
```

With visual audit:

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Document.docx" --audit
```

Visual audit only (skip text revision):

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Document.docx" --audit-only
```

### Manual mode (no API credits required)

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Document.docx" --manual
```

This creates copy-and-paste prompts in:

```
output/<document_name>/<document_name>_manual_prompts/
```

Paste those prompts into Claude, ChatGPT, or Gemini in the browser.

## Output structure

```
output/
  Document/
    Document_revised_pt_br.docx
    Document_revised_pt_br.pdf
    Document_revised_en_us.docx
    Document_revised_en_us.pdf
    Document_visual_audit.docx    (with --audit or --audit-only)
    Document_visual_audit.pdf
    _work/                        (LaTeX sources and extracted figures)
```

## Test files

```powershell
python run_test.py          # API mode with audit
python run_test_manual.py   # Manual mode, no API required
```

## Notes

- The `.env` file is listed in `.gitignore`. Do not commit API keys.
- The `output/` directory is also excluded from git.
- Embedded Word tables and figures are preserved via a placeholder system during AI revision and reinserted into all output formats.
- The tool does not guarantee results from any AI-detection system and must not be used to deceive readers.

## Changelog

### v2 (2026)
- Fixed `latex_escape`: backslash placeholder prevents curly-brace double-escaping
- Fixed `is_heading_line`: stricter heuristics reduce false positives (80-char limit, multi-sentence check, lowercase check)
- Fixed heading hierarchy in LaTeX: numbered headings now map to `\section`, `\subsection`, `\subsubsection`
- Fixed `call_claude`: accepts optional `system` parameter; visual audit now uses correct system role
- Fixed DOCX table style: replaced "Table Grid" (all borders) with three-line academic style
- Fixed ReportLab table style: replaced `GRID` with `LINEABOVE`/`LINEBELOW` three-line style
- Fixed `harmonize`: chunked for large documents to avoid context-window overflow
- Fixed LaTeX `colspec`: content-aware column widths instead of uniform width
- Added `\subsubsection` support in LaTeX and second-level headings in ReportLab PDF
- Added `.python-version` (3.11)
- Tightened `requirements.txt` with upper-bound version constraints
