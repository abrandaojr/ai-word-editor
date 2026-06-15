# AI Word Editor

[![Python](https://img.shields.io/badge/python-3.11%2B-blue.svg)](https://www.python.org/)
[![CI](https://github.com/abrandaojr/ai-word-editor/actions/workflows/ci.yml/badge.svg)](https://github.com/abrandaojr/ai-word-editor/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

AI Word Editor is a Python tool for revising Microsoft Word documents into clear,
publication-ready Brazilian Portuguese and American English. It is designed for
technical reports, policy briefs, academic manuscripts, and applied research
documents that need to remain accurate while becoming easier to read.

The tool reads a `.docx` file and produces revised Word and PDF outputs, while
preserving embedded tables, figures, maps, and placeholders.

## What It Produces

- Revised Brazilian Portuguese `.docx` and `.pdf`
- Revised American English `.docx` and `.pdf`
- LaTeX sources used to build the PDFs
- Extracted figures and tables used during PDF generation
- Optional visual communication audit for embedded figures, tables, charts, and maps
- Manual copy-and-paste prompts when API usage is not desired

## Why This Project Exists

Research documents often need two things at once: technical precision and public
readability. AI Word Editor helps with that middle ground. It keeps factual
claims, numbers, dates, institutions, locations, caveats, and document structure
intact, while improving flow, clarity, paragraph logic, and bilingual presentation.

The editing prompt follows well-established principles from academic writing,
science communication, explanatory journalism, and data visualization. Those
principles are applied silently, without adding references to the output.

## Requirements

- Python 3.11 or newer
- Anthropic API key for automatic revision
- Microsoft Word `.docx` input file

Optional:

- `pdflatex` from MiKTeX or TeX Live for LaTeX-based PDF generation
- If `pdflatex` is unavailable, the tool automatically falls back to ReportLab

## Quick Start

```powershell
git clone https://github.com/abrandaojr/ai-word-editor.git
cd ai-word-editor

python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

copy .env.example .env
notepad .env
```

Add your Anthropic API key to `.env`:

```env
ANTHROPIC_API_KEY=sk-ant-api03-your-real-key-here
```

Run the editor:

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Document.docx"
```

## Usage

Revise a document in Brazilian Portuguese and American English:

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Document.docx"
```

Run the revision and add a visual communication audit:

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Document.docx" --audit
```

Generate only the visual communication audit:

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Document.docx" --audit-only
```

Generate manual prompts without using the API:

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Document.docx" --manual
```

Manual mode writes copy-and-paste prompts to:

```text
output/<document_name>/<document_name>_manual_prompts/
```

## Output Layout

```text
output/
  Document/
    Document_revised_pt_br.docx
    Document_revised_pt_br.pdf
    Document_revised_en_us.docx
    Document_revised_en_us.pdf
    Document_visual_audit.docx
    Document_visual_audit.pdf
    _work/
      pt_br/
      en_us/
      figures/
```

## Configuration

The tool reads configuration from environment variables or `.env`:

| Variable | Default | Purpose |
| --- | --- | --- |
| `ANTHROPIC_API_KEY` | required | API key used for automatic revision |
| `ANTHROPIC_MODEL` | `claude-sonnet-4-6` | Anthropic model name |
| `MAX_CHUNK_CHARS` | `22000` | Maximum chunk size for long documents |
| `MAX_OUTPUT_TOKENS` | `20000` | Maximum output tokens per model call |
| `TEMPERATURE` | `0.2` | Revision temperature |
| `OUTPUT_DIR` | `output` | Directory for generated files |

## Helper Scripts

```powershell
python run_test.py          # API mode with visual audit
python run_test_manual.py   # Manual mode, no API required
python run_test_no_api.py   # Alias for manual mode
```

The helper scripts expect a sample file named:

```text
C:\Users\<you>\Downloads\Desmatamento na Amazônia.docx
```

Edit the filename inside the scripts if your test document has another name.

## Quality Checks

```powershell
python -m compileall -q ai_word_editor.py run_test.py run_test_manual.py run_test_no_api.py
python ai_word_editor.py --help
```

The GitHub Actions workflow runs these checks on every push and pull request.

## Notes

- `.env` is ignored by git. Do not commit API keys.
- `output/` is ignored by git because it contains generated documents.
- Embedded Word tables and figures are preserved with stable placeholders during revision and reinserted into the outputs.
- The tool is intended for editorial clarity and document production, not for deceiving readers or bypassing authorship policies.

## License

This project is released under the MIT License.
