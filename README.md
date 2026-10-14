# pdf2key

Turn any PDF into an Apple Keynote presentation that looks the same and has **editable text**.

Every heading, paragraph and label becomes a real Keynote text box that you can retype or restyle. Photos, shapes, gradients and logos look exactly as they do in the PDF. After converting, pdf2key exports every slide back out of Keynote, compares it pixel by pixel with the PDF page, and shows you how close it is.

## Requirements

- macOS with **Keynote** installed (Keynote does the final `.pptx` → `.key` step)
- Python 3.10+

## Install

```bash
git clone https://github.com/prakharg01/pdf2key.git
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
| `--separate-images` | Make each picture its own movable object (less faithful, see below) |
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

Each PDF page becomes one slide with two layers:

| Layer | Editable? | Details |
| --- | --- | --- |
| Art | Picture | Everything except text, rendered at 4K exactly as the PDF draws it. Overlapping, cropped, masked and semi-transparent images and shapes all stay correct. |
| Text | ✅ | Real text boxes. They keep the PDF's line breaks and size, colour and opacity. Baselines, letter-spacing, weight, italics and rotation are also matched (see below). |

How the text is rebuilt:

- **Baselines.** With exact line spacing *L*, Keynote puts each line's baseline 0.8 × *L* below the top of the line, whatever the font. pdf2key works out each line's spacing so every baseline lands exactly where the PDF has it.
- **Words and letter-spacing.** Words and letter-spacing are rebuilt from where each glyph sits. This matters for PDFs from browsers (Chrome, "Save as PDF"), which often draw spaced-out labels with no space characters at all.
- **Fonts with no name.** Browser PDFs often embed fonts with no name (Type3). pdf2key works out their style from the glyphs themselves:
  - **Weight:** each font's ink is compared with the document's body font to sort it into *regular*, *medium* or *bold*.
  - **Italics:** a deslant test finds the shear angle that makes the vertical strokes sharpest.
  - **Monospace:** a font counts as monospace if all its glyphs have the same advance.
- **Unrelated text.** Text that only happens to share a PDF line (such as a figure and its caption) is split into separate boxes.

Pages that are a different size from the first page are scaled to fit the slide.

### `--separate-images`

This mode places each embedded image as its own picture, on top of a background colour and under a transparent line-art overlay, so you can drag photos around. It can't reproduce PDF stacking, clipping or opacity, so use it only for simple documents.

## Getting the closest match

- **Install the PDF's fonts.** Fonts embedded in a PDF are usually partial subsets that Keynote can't use. If a font isn't installed on your Mac, pdf2key picks a lookalike (Helvetica Neue, Menlo or Times New Roman) and scales it so line widths still match. The app lists which fonts it substituted.
- **Scanned PDFs** contain only images, with no text to extract. Run OCR first (for example `ocrmypdf`).

## Development

```bash
.venv/bin/pip install -e ".[dev]"
.venv/bin/pytest
```

The tests cover the PDF → `.pptx` stage and don't need Keynote. Put real PDFs to try in `tests/pdfs/`. PDFs and generated `.key`/`.pptx` files are git-ignored, so documents never get committed.

```
src/pdf2key/
  converter.py   conversion engine and accuracy check
  cli.py         `pdf2key` command
  web.py         `pdf2key-web` Flask app
  static/        web UI
tests/
  test_converter.py
  pdfs/          local sample PDFs (ignored by git)
```
