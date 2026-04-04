# officecat

A CLI tool to view Office files in the terminal. Think `cat` but for `.docx`, `.pptx`, `.xlsx`, and `.csv` files.

## Installation

```bash
pip install -e .
```

## Usage

```bash
officecat report.docx              # interactive TUI viewer
officecat report.docx --rich       # colored formatted dump (like cat)
officecat budget.xlsx | head       # plain text (auto-detected pipe)
officecat slides.pptx --json       # JSON output
```

### Output Modes

- **TUI** (default when TTY): Interactive viewer with scrolling, tabs for sheets/slides, and keyboard navigation.
- **Rich** (`--rich`): Colored, formatted dump to stdout. Non-interactive, pairs well with `less -R`.
- **Plain** (auto when piped, or `--plain`): Pipe-friendly text output.
- **JSON** (`--json`): Structured output for scripting.

### Options

| Flag | Short | Description |
|---|---|---|
| `--rich` | `-r` | Colored formatted dump (non-interactive) |
| `--plain` | `-p` | Plain text output, no colors |
| `--json` | `-j` | JSON output |
| `--head N` | `-n N` | Show first N items (rows/paragraphs/slides) |
| `--list` | `-l` | List contents only (sheet names, slide titles, heading outline) |
| `--sheet S` | `-s S` | Select sheet by name or 1-based index (xlsx/csv only) |
| `--slide N` | | Show only slide N (pptx only) |
| `--headers N` | `-h N` | Promote row N as column headers (xlsx/csv, default: 1, 0 to disable) |
| `--all` | `-a` | Disable the default 500-row cap for large spreadsheets |

### TUI Key Bindings

| Key | Action |
|---|---|
| `q` / `Esc` | Quit |
| `Up` / `Down` / `j` / `k` | Scroll |
| `PgUp` / `PgDn` | Page scroll |
| `Home` / `End` | Jump to top/bottom |
| `Tab` / `Shift+Tab` | Next/previous sheet or slide |

### Examples

```bash
# View first 10 rows of an Excel file in TUI
officecat budget.xlsx --head 10

# View a specific sheet
officecat budget.xlsx --sheet "Q4 Summary"

# List all sheet names
officecat budget.xlsx --list

# View slide 3 of a presentation
officecat deck.pptx --slide 3

# Colored dump piped to less
officecat report.docx --rich | less -R

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
