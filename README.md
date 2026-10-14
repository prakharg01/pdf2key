# pdf2key

Turn any PDF into an **editable** Apple Keynote presentation.

You get real Keynote text boxes and pictures, not a deck of page screenshots. You can retype a heading, move a photo, or restyle a paragraph, and the layout stays the same as the PDF.

After converting, pdf2key exports every slide back out of Keynote, compares it pixel by pixel with the PDF page, and shows you how close it is.

## Requirements

- macOS with **Keynote** installed (Keynote does the final `.pptx` → `.key` step)
- Python 3.10+

## Install

```bash
git clone https://github.com/<you>/pdf2key.git
cd pdf2key
python3 -m venv .venv
.venv/bin/pip install -e .
```

## Use it

### Web app (upload and download)

```bash
.venv/bin/pdf2key-web
```

This opens <http://127.0.0.1:5001>. Drop in a PDF and download the `.key`. For each slide you see the PDF, the Keynote result, and a difference heatmap. Everything runs locally and nothing leaves your Mac.

### Command line

```bash
.venv/bin/pdf2key deck.pdf              # writes deck.key next to the PDF
.venv/bin/pdf2key deck.pdf out.key
.venv/bin/pdf2key deck.pdf --pptx-only  # stop at .pptx (works without Keynote)
```

| Option | What it does |
| --- | --- |
| `--no-background` | Leave out the flattened line-art layer |
| `--no-verify` | Skip the PDF-vs-Keynote comparison (faster) |
| `--pptx-only` | Write a `.pptx` instead of a `.key` |

### From Python

```python
from pdf2key import convert

reports = convert("deck.pdf", "deck.key")
for r in reports:
    print(r.page, r.match, r.off_pixels)
```

## How it works

```
PDF ──PyMuPDF──▶ .pptx (python-pptx) ──Keynote via AppleScript──▶ .key ──▶ accuracy check
```

Each PDF page becomes one slide, built in layers from the bottom up:

| Layer | Editable? | Details |
| --- | --- | --- |
| Background | ✅ | A full-page fill becomes the slide background colour |
| Pictures | ✅ | Embedded images, at original quality, with transparency kept |
| Line art | ❌ | Vector lines, shapes and gradients, flattened into one 300 dpi transparent image |
| Text | ✅ | Real text boxes with the PDF's line breaks, line spacing, size, colour, bold/italic and rotation |

Pages that are a different size from the first page are scaled to fit the slide.

## Getting the closest match

- **Install the PDF's fonts.** Fonts embedded in a PDF are usually partial subsets that Keynote can't use. If a font isn't installed on your Mac, pdf2key picks a mono, serif or sans lookalike and scales it so line widths still match. The app lists which fonts it substituted. Install those fonts and convert again for a near-exact result.
- **Scanned PDFs** contain only images, with no text to extract. Run OCR first (for example `ocrmypdf`).
- **Vector graphics** such as charts, icons and diagram lines come through as one flattened image, so they look right but you can't edit them shape by shape.

## Development

```bash
.venv/bin/pip install -e ".[dev]"
.venv/bin/pytest
```

The tests cover the PDF → `.pptx` stage and don't need Keynote.

```
src/pdf2key/
  converter.py   conversion engine and accuracy check
  cli.py         `pdf2key` command
  web.py         `pdf2key-web` Flask app
  static/        web UI
tests/
```
