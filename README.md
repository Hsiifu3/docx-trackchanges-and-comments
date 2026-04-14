# docx-trackchanges-and-comments

A lightweight Python utility for applying native Microsoft Word Track Changes to `.docx` files with minimal OOXML modifications.

## Features

- Paragraph replacement with tracked insertion/deletion
- Inline text replacement with tracked changes
- Replace only the nth inline match
- Paragraph deletion
- Insert paragraph after anchor paragraph
- Add Word comments
- Replace text in headers and footers
- Support body paragraphs and table cell paragraphs

## File layout

- `scripts/track_changes.py` — main CLI script
- `SKILL.md` — usage-oriented skill documentation

## Requirements

- Python 3.10+
- `lxml`

Install dependency:

```bash
pip install lxml
```

## Quick start

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace-inline "old phrase" "new phrase"
```

## Examples

Replace a whole paragraph:

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace "Old paragraph" "New paragraph"
```

Replace the 2nd inline match:

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace-inline-nth "dynamic response" "vibration response" 2
```

Add a comment:

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Reviewer" \
  --comment "finite element" "Please add software version here"
```

## Notes

This repository intentionally excludes sample DOCX files and local metadata so it is safe to publish as a clean public repository.

## License

MIT
