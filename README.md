# AI Word Editor

Python tool to revise a Microsoft Word document for a mixed audience and generate formatted outputs in Word, LaTeX, and PDF.

The tool reads one `.docx` file as input and produces:

- a revised Word document in Brazilian Portuguese
- a revised Word document in American English
- a formatted PDF in Brazilian Portuguese
- a formatted PDF in American English
- LaTeX source files used to build the PDFs
- extracted figures and detected tables from the input Word file
- a visual communication audit in Markdown, Word, and PDF

The visual audit checks figures, maps, charts, and tables using structured criteria for chart type, data integrity, titles, labels, color, annotations, captions, cross-references, and overall visual coherence.

The PDF workflow is based on the logic of `latex-doc-builder`: Word content is converted to structured text, figures and tables are identified, LaTeX is assembled, and `pdflatex` is used when available. If LaTeX is not installed, the tool falls back to ReportLab and still generates PDFs.

## What it detects from the Word file

The script automatically scans the input `.docx` and detects:

- normal paragraphs
- headings
- embedded Word tables
- embedded figures/images
- likely figure and table captions when they appear next to the object
- surrounding text that refers to each figure or table

Tables and figures are protected with placeholders during AI revision, such as:

```text
[[TABLE:table_001]]
[[FIGURE:figure_001]]
```

The model is instructed not to change those placeholders. After revision, the script reinserts the original tables and figures into the Word and PDF outputs.

## Editorial guidance

The prompt uses principles from:

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

The prompt also includes rules to avoid robotic prose: generic openings, repetitive transitions, inflated adjectives, mechanical paragraph rhythm, and stock summaries. The goal is natural, careful, human-sounding editing while preserving factual accuracy and the author's voice. The tool does not guarantee results from any AI-detection system and should not be used for deception.

## Installation

Windows PowerShell:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env
notepad .env
```

Add your real Anthropic API key to `.env`:

```text
ANTHROPIC_API_KEY=sk-ant-api03-your-real-key-here
ANTHROPIC_MODEL=claude-sonnet-4-6
MAX_CHUNK_CHARS=22000
MAX_OUTPUT_TOKENS=20000
TEMPERATURE=0.2
OUTPUT_DIR=output
```

## Optional LaTeX installation

The script can generate PDFs without LaTeX by using ReportLab. For better LaTeX-formatted PDFs, install LaTeX.

Windows:

```powershell
winget install MiKTeX.MiKTeX
```

macOS:

```bash
brew install --cask mactex
```

Linux:

```bash
sudo apt-get install texlive-full
```

## Visual communication audit

By default, API mode also produces a visual audit for each language. The audit is saved as `.md`, `.docx`, and `.pdf` in:

```text
output/<input_name>/visual_audit/
```

The audit follows rules for scientific data visualization and public-facing graphics. It reviews each detected table, figure, chart, or map and provides:

- a diagnosis of what the visual communicates
- whether the chart or map type is appropriate
- data integrity risks, including misleading axes, dual axes, unnormalized choropleths, 3D effects, pie charts with too many slices, and poor color choices
- recommendations for declarative titles, subtitles, labels, source lines, annotations, and captions
- map-specific guidance on projection, scale bar, north arrow, locator inset, class breaks, and basemap discipline
- table-specific guidance using APA-style three-line tables
- stronger cross-references in the surrounding text
- a document-level visual coherence summary

To skip the audit:

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Desmatamento na Amazônia.docx" --skip-visual-audit
```

## API mode

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Desmatamento na Amazônia.docx"
```

Expected outputs:

```text
output/<input_name>/docx/<input_name>_revised_ptbr.docx
output/<input_name>/docx/<input_name>_revised_enus.docx
output/<input_name>/pdf/<input_name>_revised_ptbr.pdf
output/<input_name>/pdf/<input_name>_revised_enus.pdf
output/<input_name>/latex_build/*.tex
output/<input_name>/work/figures/*
output/<input_name>/input_structure_report.md
output/<input_name>/visual_audit/<input_name>_visual_audit_ptbr.md
output/<input_name>/visual_audit/<input_name>_visual_audit_ptbr.docx
output/<input_name>/visual_audit/<input_name>_visual_audit_ptbr.pdf
output/<input_name>/visual_audit/<input_name>_visual_audit_enus.md
output/<input_name>/visual_audit/<input_name>_visual_audit_enus.docx
output/<input_name>/visual_audit/<input_name>_visual_audit_enus.pdf
```

## Manual mode, no API call

If you do not have API credits, generate copy-and-paste prompts:

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Desmatamento na Amazônia.docx" --manual
```

Prompts are saved in:

```text
output/<input_name>/manual_prompts/
```

## PDF/layout test without API

To test only the Word extraction and PDF generation pipeline, without Claude:

```powershell
python ai_word_editor.py "C:\Users\Amintas\Downloads\Desmatamento na Amazônia.docx" --skip-ai
```

This uses the original Word text and still extracts figures/tables and creates DOCX/PDF outputs. In this mode, the visual audit is an inventory placeholder rather than a full AI critique.

## Test helpers

For the test Word file in your Downloads folder:

```powershell
python run_test.py
```

Manual prompt test:

```powershell
python run_test_manual.py
```

No-API PDF layout test:

```powershell
python run_test_no_api.py
```

## Update GitHub repository

```powershell
cd $HOME\Downloads
Expand-Archive ai-word-editor-visual.zip -DestinationPath . -Force
cd ai-word-editor-visual

git init
git add .
git commit -m "Add PDF outputs and visual communication audit"
git branch -M main
git remote add origin https://github.com/abrandaojr/ai-word-editor.git
git push -u origin main --force
```

If the remote already exists:

```powershell
git remote set-url origin https://github.com/abrandaojr/ai-word-editor.git
git push -u origin main --force
```

## Important notes

Do not upload `.env` to GitHub. It is ignored by `.gitignore`.

PDF quality is best with LaTeX installed. Without LaTeX, ReportLab fallback PDFs are still generated but are simpler.
