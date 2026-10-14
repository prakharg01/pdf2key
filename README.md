# pdf2key

Turn a **PDF** or an **HTML deck export** (such as a Claude Design `.zip`) into an Apple Keynote presentation. The result looks the same as the original, and you can edit it.

| You give it | You get |
| --- | --- |
| A PDF | Editable text. Everything else (photos, shapes, gradients) is one picture behind the text, exactly as the PDF draws it. |
| An HTML deck export (`.zip`, unzipped folder, or its `.html`) *(experimental)* | Editable text **and movable photos**. Each photo is its own Keynote picture, made from the original image file with its crop. Cards, overlays and backgrounds are transparent layers in between. |

After converting, pdf2key exports every slide back out of Keynote, compares it pixel by pixel with the original, and gives you a **conversion report**:

- **Results:** slides, time taken, the average match and the lowest-matching slides.
- **Objects created:** text boxes, photos, shapes, overlays and photo groups.
- **Font changes:** every font recognised from glyphs, used as named, substituted or estimated, with character counts.
- **Text adjustments:** letter-spacing, justification, rotated text, and boxes resized for substitute fonts.
- **Warnings:** for example, images that weren't found.

The web app shows the report as a card. The command line prints it at the end.

## Requirements

- macOS with **Keynote** installed (Keynote does the final `.pptx` → `.key` step)
- Python 3.10+
- For HTML decks: **Google Chrome**, plus the `html` extra. It installs Playwright, which drives your installed Chrome (no extra browser download), and SciPy, which splits layers into objects.

## Install

```bash
git clone https://github.com/prakharg01/pdf2key.git
cd pdf2key
python3 -m venv .venv
.venv/bin/pip install -e ".[html]"   # or just -e . if you only convert PDFs
```

## Use it

### Web app (upload and download)

```bash
.venv/bin/pdf2key-web
```

This opens <http://127.0.0.1:5001>. Drop in a **PDF** or an **HTML deck export `.zip`** and download the `.key`. For each slide you see the original, the Keynote result, and a difference heatmap. Everything runs locally and nothing leaves your Mac.

### Command line

```bash
# PDF
.venv/bin/pdf2key deck.pdf                       # writes deck.key next to the PDF
.venv/bin/pdf2key deck.pdf out.key

# HTML deck export: the zip, the unzipped folder, or the .html inside it
.venv/bin/pdf2key "Short Deck.zip" out.key
.venv/bin/pdf2key "export/KinesthetIQ Short Deck" out.key
.venv/bin/pdf2key "export/KinesthetIQ Short Deck/KinesthetIQ Short Deck.html" out.key
```

| Option | What it does |
| --- | --- |
| `--no-verify` | Skip the comparison with the original (faster) |
| `--pptx-only` | PDF only: write a `.pptx` instead of a `.key` (works without Keynote) |
| `--separate-images` | PDF only: make each embedded image movable (less faithful, see below) |

### From Python

```python
from pdf2key import convert
from pdf2key.html import convert_html

reports = convert("deck.pdf", "deck.key")
reports = convert_html("Short Deck.zip", "deck.key")
for r in reports:
    print(r.page, r.match, r.off_pixels)
```

## HTML deck exports: what to pass in

An HTML deck isn't just the `.html` file. The page refers to its photos (`assets/…`), styles and font links (`_ds/…`), and scripts **by relative path**. pdf2key needs those files next to the HTML, as they are in the export:

```
KinesthetIQ Short Deck.zip
└── export/KinesthetIQ Short Deck/
    ├── KinesthetIQ Short Deck.html   ← the deck
    ├── assets/                       ← photos and logos
    ├── _ds/                          ← design-system styles and fonts
    └── deck-stage.js, support.js
```

So any of these works:

- **The `.zip` as exported.** This is the simplest, and the only option in the web app.
- **The unzipped folder.** pdf2key picks the deck's `.html` (the largest one) inside it.
- **The `.html` file**, as long as it's still inside the unzipped folder.

An `.html` copied out on its own converts with its text but without photos, and with fallback styling.

Each top-level `<section>` is one slide, at the size the deck declares (1920 × 1080 for Claude Design exports). A page with no sections becomes a single slide. Install the deck's fonts for the closest text match (see below).

## How it works

### PDFs

```
PDF ──PyMuPDF──▶ .pptx (python-pptx) ──Keynote via AppleScript──▶ .key ──▶ accuracy check
```

Each PDF page becomes one slide with two layers:

| Layer | Editable? | Details |
| --- | --- | --- |
| Art | Picture | Everything except text, rendered at 4K exactly as the PDF draws it. Overlapping, cropped, masked and semi-transparent images and shapes all stay correct. |
| Text | ✅ | Real text boxes. They keep the PDF's line breaks and size, colour and opacity. Baselines, letter-spacing, weight, italics and rotation are also matched (see below). |

Pages that are a different size from the first page are scaled to fit the slide.

### HTML decks (experimental)

A PDF only records "draw this image here". The HTML knows which element is a photo, how it's cropped, and what's layered over it:

1. **Layout.** Chrome lays out each slide at the deck's size with the deck's fonts.
2. **Photos.** Every element is listed in paint order. Each `<img>` or CSS background image becomes a Keynote picture made from the **original file**. It keeps the crop the browser applies (`object-fit`, `object-position`, `background-size`, overflow clipping), along with opacity and rounded corners. Pictures with effects Keynote can't reproduce (filters, masks, rotation, blend modes) are captured on their own, still as separate pictures.
3. **Layers.** Decorations between two pictures (backgrounds, cards, gradients, borders, SVG arrows) become transparent layers in the right stacking order. Each layer is **split into its separate islands**, so every card, arrow, badge and overlay is its own object. Paint that covers the whole slide (the background, a full-slide tint) is kept as its own layer, so it never fuses the shapes above it.
4. **Groups and names.** A photo is grouped with any overlay that lies only on it (a card's gradient, a caption plate), so they move together. Double-click the group to reach the photo and crop or replace it. Every object is named ("Photo · robot.png", "Overlay", "Shape", "Background"), so you can pick anything from Keynote's object list (**View › Show Object List**), even under a full-slide tint.
5. **Text.** Chrome also prints the slides to PDF, and the text pipeline below puts editable text on top.

On the sample 25-slide deck this gives 123 photos, 275 shapes, 15 overlays and 7 photo groups as separate objects, with a 97.6% average match. Cards, borders and gradients are pictures, not native Keynote shapes you can recolour.

### How the text is rebuilt (both inputs)

- **Baselines.** With exact line spacing *L*, Keynote puts each line's baseline 0.8 × *L* below the top of the line, whatever the font. pdf2key works out each line's spacing so every baseline lands exactly where the original has it.
- **Words, letter-spacing and justification.** These are rebuilt from where each glyph sits. This matters for browser-made PDFs, which often draw spaced-out labels with no space characters. Justified lines keep their widths because the extra word spacing goes onto the spaces.
- **Fonts with no name.** Browser PDFs often embed fonts with no name (Type3). pdf2key **recognises them from their glyph widths**. Every face has its own pattern of letter widths, so it compares the distance between neighbouring glyphs with every font installed on your Mac and picks the face that matches, such as *Montserrat ExtraBold* or *Inter SemiBold Italic*. Variable fonts are checked at each of their named weights. The same comparison also gives the exact letter-spacing. Where several faces have the same widths (all monospace fonts, or an italic and its upright), it draws the glyphs in each candidate and keeps the one that looks most like the original. The list of installed fonts is cached in `~/Library/Caches/pdf2key`.
- **Fallback when no installed font matches.** pdf2key estimates the style from the glyphs instead:
  - **Weight:** each font's ink is compared with the document's body font to sort it into *regular*, *medium* or *bold*.
  - **Italics:** a deslant test finds the shear angle that makes the vertical strokes sharpest.
  - **Monospace:** a font counts as monospace if all its glyphs have the same advance.

  It then uses Helvetica Neue or Menlo at that style.
- **Unrelated text.** Text that only happens to share a line (such as a figure and its caption) is split into separate boxes.

### `--separate-images` (PDF only)

This mode places each embedded PDF image as its own picture, on top of a background colour and under a transparent line-art overlay. It can't reproduce PDF stacking, clipping or opacity, so use it only for simple documents. For decks, use the HTML export instead.

## Getting the closest match

- **Install the fonts.** Fonts embedded in a PDF are usually partial subsets that Keynote can't use. If a font isn't installed on your Mac, pdf2key picks a lookalike (Helvetica Neue, Menlo or Times New Roman) and scales it so line widths still match. The app lists the fonts it substituted and the ones it recognised. Free fonts install with Homebrew, for example `brew install --cask font-inter font-montserrat font-jetbrains-mono`. **Restart Keynote** after installing fonts, because it only picks up new fonts when it launches.
- **Scanned PDFs** contain only images, with no text to extract. Run OCR first (for example `ocrmypdf`).

## Development

```bash
.venv/bin/pip install -e ".[dev,html]"
.venv/bin/pytest
```

The tests cover building the `.pptx` from PDFs and HTML and don't need Keynote. The HTML tests skip themselves if Chrome isn't available. Put real decks to try in `tests/pdfs/`. Everything there except the placeholder is git-ignored, so documents never get committed.

```
src/pdf2key/
  converter.py   PDF engine, Keynote step and accuracy check
  fonts.py       installed-font catalog and glyph-width font matching
  html.py        HTML deck input (movable pictures), experimental
  report.py      the conversion report
  cli.py         `pdf2key` command
  web.py         `pdf2key-web` Flask app
  static/        web UI
tests/
  test_converter.py
  test_html.py
  pdfs/          local sample decks (ignored by git)
```
