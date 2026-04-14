# docx-trackchanges-and-comments

A lightweight Python utility for applying **native Microsoft Word Track Changes** and **comments** to `.docx` files with minimal OOXML modifications.

This project is designed for revision-heavy workflows such as paper editing, supervisor feedback, collaborative manuscript polishing, and tracked DOCX generation.

## Highlights

- Preserve **Word-native review markup** instead of plain text diffs
- Apply tracked changes directly to `.docx` files
- Keep changes as small as possible by modifying only the necessary OOXML parts
- Support comments, inline replacements, paragraph operations, and header/footer text replacement
- Work on body paragraphs **and table cell paragraphs**

## Features

### Tracked revision operations
- Replace an entire paragraph with tracked deletion + insertion
- Replace inline text with tracked changes
- Replace only the **nth** inline match
- Delete a paragraph with revision marks
- Insert a new paragraph after an anchor paragraph

### Review helpers
- Add Word comments to matching text
- Replace text in headers and footers
- Emit replacement logs for easier verification

### Scope currently supported
- Main body paragraphs
- Table cell paragraphs
- Header/footer paragraphs

## Repository layout

- `scripts/track_changes.py` — main CLI tool
- `SKILL.md` — detailed usage guide and examples

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

## CLI examples

### 1. Replace a whole paragraph

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace "Old paragraph" "New paragraph"
```

### 2. Replace inline text

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace-inline "better prediction accuracy" "higher prediction accuracy"
```

### 3. Replace only the 2nd inline match

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace-inline-nth "dynamic response" "vibration response" 2
```

### 4. Delete a paragraph

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --delete "Paragraph to remove"
```

### 5. Insert a paragraph after an anchor

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --insert-after "Anchor paragraph" "Inserted paragraph"
```

### 6. Add comments

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Reviewer" \
  --comment "finite element" "Please add the software version here" \
  --comment "dynamic response" "Consider citing a recent reference"
```

### 7. Replace text in headers and footers

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace-hf "Company Name" "New Company" \
  --replace-hf "2023" "2024"
```

### 8. Run multiple actions in one pass

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace "Old paragraph A" "New paragraph A" \
  --replace-inline "old phrase" "new phrase" \
  --replace-inline-nth "dynamic response" "vibration response" 2 \
  --comment "finite element" "Please add a reference" \
  --replace-hf "Old Institution" "New Institution" \
  --delete "Old paragraph B" \
  --insert-after "Paragraph C" "Inserted paragraph D"
```

## How it works

The script follows a **minimal-intrusion OOXML** strategy:

- unpack the `.docx`
- modify the required XML parts
- write Word-compatible tracked insertions/deletions/comments
- repackage the document

The goal is to preserve the original Word shell as much as possible while still producing native review markup visible in Microsoft Word.

## Typical use cases

- Academic paper revision with visible track changes
- Supervisor/student collaborative editing
- Reviewer-style comment injection into DOCX manuscripts
- Batch-like manuscript cleanup with controlled tracked edits
- Preparing revision-friendly files for coauthors

## Current limitations

- `--replace` and `--delete` require **exact paragraph matches**
- Inline replacement supports both **single-run** and **cross-run** matching, but complex formatting may still be imperfect
- If target text crosses fragile structures such as some field-code boundaries, the script may skip replacement to avoid corrupting the document
- Text boxes, footnotes, and endnotes are **not yet supported**
- For complex mixed-style runs, formatting is preserved on a best-effort basis, usually using the first relevant run style

## Safety / publishing note

This repository intentionally excludes:

- sample `.docx` files
- local test artifacts
- local publishing metadata
- personal document metadata from example files

That keeps the public repository cleaner and safer to share.

## Roadmap ideas

- Support text boxes, footnotes, and endnotes
- Better style preservation for complex cross-run replacements
- Native automation for accepting/rejecting revisions
- Cleaner packaging as a reusable Python tool or pip package

## Documentation

For more detailed examples and Chinese usage notes, see:

- `SKILL.md`

## License

MIT
