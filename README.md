# officecat

A CLI tool to view Office files in the terminal. Think `cat` but for `.docx`, `.pptx`, `.xlsx`, and `.csv` files.

## Installation

```bash
pip install -e .
```

## Usage

```bash
officecat report.docx
officecat slides.pptx
officecat budget.xlsx
officecat data.csv
```

### Options

| Flag | Short | Description |
|---|---|---|
| `--plain` | `-p` | Plain text output, no colors/boxes |
| `--json` | `-j` | JSON output |
| `--head N` | `-n N` | Show first N items (rows/paragraphs/slides) |
| `--list` | `-l` | List contents only (sheet names, slide titles, heading outline) |
| `--sheet S` | `-s S` | Select sheet by name or 1-based index (xlsx only) |
| `--slide N` | | Show only slide N (pptx only) |
| `--headers N` | `-h N` | Promote row N as column headers (xlsx/csv, default: 1, 0 to disable) |
| `--all` | `-a` | Disable the default 500-row cap for large spreadsheets |

### Output Modes

- **Rich** (default): Colored, formatted terminal output with tables and panels.
- **Plain**: Pipe-friendly text. Auto-selected when stdout is not a TTY.
- **JSON**: Structured output via `--json`.

### Examples

```bash
# View first 10 rows of an Excel file
officecat budget.xlsx --head 10

# View a specific sheet
officecat budget.xlsx --sheet "Q4 Summary"

# List all sheet names
officecat budget.xlsx --list

# View slide 3 of a presentation
officecat deck.pptx --slide 3

# Pipe-friendly CSV output
officecat data.xlsx --plain | grep "revenue"

# JSON output for scripting
officecat report.docx --json | jq '.blocks'
```

## Known Limitations

- DOCX tables are rendered after all paragraphs, not in their original position within the document body.
- DOCX list detection is style-name-based and may miss some numbered/bulleted lists that use custom styles.
- PPTX grouped shapes and embedded tables show as placeholders, not full content.
- PPTX charts and SmartArt are not extracted.
- XLSX formulas show cached values (last-computed), not the formula strings.
- No decryption of password-protected files.
- No OCR for scanned/image-based content.
- No tracked changes or comment rendering.
- Legacy binary formats (`.doc`, `.ppt`, `.xls`) are not supported — a conversion hint is shown instead.
