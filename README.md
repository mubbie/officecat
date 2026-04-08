# officecat 🐱

A CLI tool to view Office files in the terminal. Think `cat` but for `.docx`, `.pptx`, `.xlsx`, `.csv`, and `.tsv` files.

Every supported format is converted to markdown internally, then rendered through a single unified pipeline.

## Installation

```bash
pip install officecat
```

Or install from source:

```bash
git clone https://github.com/mubbie/officecat.git
cd officecat
pip install -e .
```

## Usage

```bash
officecat report.docx              # colored formatted output (default)
officecat budget.xlsx              # spreadsheet as markdown table
officecat slides.pptx              # presentation content
officecat data.csv                 # CSV and TSV
officecat report.docx --tui        # interactive full-screen viewer
officecat budget.xlsx | head       # plain text (auto-detected pipe)
officecat slides.pptx --json       # JSON output
```

### Output Modes

- **Rich** (default): Colored, formatted output to stdout. Works with `less -R`.
- **TUI** (`--tui`): Full-screen interactive viewer with scrolling.
- **Plain** (auto when piped, or `--plain`): Raw markdown for piping to `grep`, `head`, `awk`.
- **JSON** (`--json`): `{"source": "...", "markdown": "..."}` for scripting.

### Options

| Flag | Short | Description |
|---|---|---|
| `--tui` | `-t` | Interactive full-screen viewer |
| `--plain` | `-p` | Raw markdown text, no colors |
| `--json` | `-j` | JSON output |
| `--head N` | `-n N` | Show first N lines |
| `--sheet S` | `-s S` | Select sheet by name or 1-based index (xlsx only) |
| `--slide N` | | Show only slide N (pptx only) |
| `--headers N` | `-h N` | Promote row N as headers (xlsx/csv, default: 1, 0 to disable) |
| `--all` | `-a` | Disable the default 500-row cap |

### TUI Key Bindings

| Key | Action |
|---|---|
| `q` | Quit |
| `Up` / `Down` | Scroll |
| `PgUp` / `PgDn` | Page scroll |
| `Home` / `End` | Jump to top/bottom |

### Examples

```bash
# Quick view of a document
officecat report.docx

# Browse interactively
officecat report.docx --tui

# Specific sheet
officecat budget.xlsx --sheet "Q4 Summary"

# Specific slide
officecat deck.pptx --slide 3

# First 10 lines
officecat budget.xlsx --head 10

# JSON output
officecat report.docx --json | jq '.markdown'

# Pipe to grep
officecat data.xlsx --plain | grep "revenue"
```

## Supported Formats

- Word (.docx): headings, paragraphs, lists, tables in document order
- PowerPoint (.pptx): slides, shapes, images, speaker notes, hidden slides
- Excel (.xlsx): all sheets, row cap, header promotion
- CSV (.csv): auto-delimited
- TSV (.tsv): tab-delimited

Legacy binary formats (`.doc`, `.ppt`, `.xls`) show a conversion hint.

## Known Limitations

- All content is rendered as markdown. Spreadsheet tables are markdown tables, not interactive grids.
- DOCX list detection is style-name-based and may miss custom list styles.
- PPTX grouped shapes and embedded tables show as placeholders.
- PPTX charts and SmartArt are not extracted.
- XLSX formulas show cached/computed values, not formula strings.
- Large spreadsheets are capped at 500 rows by default. Use `--all` to show everything.
- TUI enforces a 1000-line cap with `--all` for performance. Use `--plain` for full output.
- No decryption of password-protected files.
- Legacy binary formats (.doc, .ppt, .xls) are not supported.
