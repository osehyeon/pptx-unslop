# pptx-unslop

A Claude Code skill for slide decks (.pptx) that do not look machine-made. It fixes one value per role (colour, type size, line weight), limits figures to PowerPoint shapes and matplotlib, removes filler text from figures, and checks the deck as PowerPoint itself renders it. The skill text is in Korean.

## Install

Clone this repository into `~/.claude/skills/pptx-unslop`.

## Requirements

macOS with Microsoft PowerPoint, poppler, Python 3 with lxml. Pillow, python-pptx and matplotlib are optional.

## Scripts

- `scripts/check_deck.py deck.pptx [--fix-lang out.pptx]` reports proofing language, bold on a Light face, embedded fonts, a logo that moves between slides, and the type sizes in use.
- `scripts/pptx_pdf.py deck.pptx --grid` renders the deck with PowerPoint to a PDF, one PNG per slide and 2x2 contact sheets.
