# officecat

A CLI tool to view Office files in the terminal. Think `cat` but for `.docx`, `.pptx`, `.xlsx`, `.csv`, and `.tsv` files.

Powered by Microsoft's [MarkItDown](https://github.com/microsoft/markitdown) for file-to-markdown conversion and [Textual](https://github.com/Textualize/textual) for the interactive TUI.

## Installation

```bash
pip install -e .
```

## Usage

```bash
officecat report.docx              # colored formatted output (default)
officecat budget.xlsx              # spreadsheet as markdown table
officecat slides.pptx              # presentation content
officecat data.csv                 # CSV and TSV too
officecat report.docx --tui        # interactive TUI with search & TOC
officecat budget.xlsx | head       # plain text (auto-detected pipe)
officecat slides.pptx --json       # JSON output
```

### Output Modes

- **Rich** (default): Colored, formatted output to stdout. Non-interactive, works with `less -R`.
- **TUI** (`--tui`): Interactive viewer with scrolling, search, table of contents sidebar, and keyboard navigation.
- **Plain** (auto when piped, or `--plain`): Raw markdown text for piping to `grep`, `head`, `awk`, etc.
- **JSON** (`--json`): Structured output for scripting — `{"source": "...", "markdown": "..."}`.

### Options

| Flag | Short | Description |
|---|---|---|
| `--tui` | `-t` | Interactive TUI viewer with search and TOC |
| `--plain` | `-p` | Raw markdown text, no colors |
| `--json` | `-j` | JSON output |
| `--head N` | `-n N` | Show first N lines of markdown output |
| `--list` | `-l` | Show file metadata and heading outline |

### TUI Key Bindings

| Key | Action |
|---|---|
| `q` / `Esc` | Quit |
| `t` | Toggle table of contents sidebar |
| `/` | Open search |
| `Enter` | Next search result |
| `Esc` (in search) | Close search |
| `Up` / `Down` | Scroll |
| `PgUp` / `PgDn` | Page scroll |
| `Home` / `End` | Jump to top/bottom |

### Examples

```bash
# Quick view of a document
officecat report.docx

# Browse interactively with search
officecat report.docx --tui

# View first 10 lines
officecat budget.xlsx --head 10

# Heading outline
officecat report.docx --list

# Colored output piped to less
officecat report.docx | less -R

# JSON output for scripting
officecat report.docx --json | jq '.markdown'

# Pipe plain markdown to grep
officecat data.xlsx --plain | grep "revenue"
```

## Supported Formats

- Word (.docx)
- PowerPoint (.pptx)
- Excel (.xlsx, .xls)
- CSV (.csv)
- TSV (.tsv)

Legacy binary formats (`.doc`, `.ppt`) are not supported — a conversion hint is shown instead.

## Known Limitations

- All content is rendered as markdown — spreadsheet tables are markdown tables, not interactive grids.
- Layout fidelity is limited — you see the content, not the original formatting or spatial arrangement.
- XLSX formulas show cached values, not formula strings.
- Password-protected/encrypted files are not supported.
- Very large files may take a moment to convert (a spinner is shown while loading).
