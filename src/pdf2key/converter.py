"""Core PDF -> editable Keynote conversion.

Each PDF page becomes a slide, layered bottom to top as:
  1. background - a full-page fill becomes the slide background color
  2. pictures   - embedded images (original quality, transparency kept)
  3. line art   - vector shapes/lines/gradients, flattened into one 300 dpi
                  transparent image (optional)
  4. text       - real text boxes with the PDF's exact line breaks, line
                  spacing, size, color, bold/italic, and rotation

Pipeline: PDF --PyMuPDF--> .pptx (python-pptx) --Keynote (AppleScript)--> .key
Afterwards every slide is exported from Keynote and compared with the PDF page.
"""
import math
import re
import statistics
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Emu, Pt

PT_TO_EMU = 12700
FONT_DIRS = [Path("/Library/Fonts"), Path("/System/Library/Fonts"), Path.home() / "Library/Fonts"]
WEIGHT_WORDS = {"thin", "extralight", "ultralight", "light", "medium", "semibold", "demibold",
                "extrabold", "ultrabold", "black", "heavy", "condensed", "narrow"}
BOLD_TOKENS = {"bold", "boldital", "bolditalic", "boldoblique", "bd", "bdit"}


# --------------------------------------------------------------------------- #
# Geometry: map PDF points onto the slide (pages of different sizes are fit).
# --------------------------------------------------------------------------- #
def _emu(v):
    return Emu(int(round(v * PT_TO_EMU)))


class _Fit:
    def __init__(self, rect, slide_w, slide_h):
        self.s = min(slide_w / rect.width, slide_h / rect.height)
        self.ox = (slide_w - rect.width * self.s) / 2
        self.oy = (slide_h - rect.height * self.s) / 2

    def x(self, v):
        return _emu(self.ox + v * self.s)

    def y(self, v):
        return _emu(self.oy + v * self.s)

    def d(self, v):
        return _emu(max(v, 0) * self.s)


# --------------------------------------------------------------------------- #
# Fonts
# --------------------------------------------------------------------------- #
_font_stems = None


def _norm(name):
    return re.sub(r"[^a-z0-9]", "", name.lower())


def _installed_stems():
    global _font_stems
    if _font_stems is None:
        _font_stems = set()
        for d in FONT_DIRS:
            if d.exists():
                for f in d.rglob("*"):
                    if f.suffix.lower() in (".ttf", ".otf", ".ttc", ".dfont"):
                        _font_stems.add(_norm(re.split(r"[-\[]", f.stem)[0]))
    return _font_stems


def _is_installed(base):
    n = re.sub(r"(mt|ps)$", "", _norm(base))
    stems = _installed_stems()
    return n in stems or (len(n) >= 4 and any(s.startswith(n) for s in stems))


def _font_info(pdf_name, flags):
    """Return (family, bold, italic, measuring_font, exact).

    exact=False means the real font isn't installed and a lookalike is used, so
    sizes are scaled per text box to keep line widths close to the PDF.
    """
    name = pdf_name.split("+")[-1]
    parts = [p for p in re.split(r"[-,]", name) if p] or ["Helvetica"]
    base, styles = parts[0], [p.lower() for p in parts[1:]]
    bold = bool(set(styles) & BOLD_TOKENS) or bool(flags & 16)
    italic = bool(flags & 2) or any("ital" in s or "oblique" in s for s in styles)
    weight = [w for w in styles if w in WEIGHT_WORDS]
    pretty = re.sub(r"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", " ", base)
    pretty = re.sub(r"\s(MT|PS)$", "", pretty)
    low = _norm(base)

    if _is_installed(base):
        family = " ".join([pretty] + [w.capitalize() for w in weight])
        return family, bold, italic, None, True
    if any(k in low for k in ("mono", "courier", "consolas", "menlo", "typewriter", "code")):
        return "Menlo", bold, italic, "cour", False
    if ("serif" in low and "sans" not in low) or any(
            k in low for k in ("times", "georgia", "garamond", "palatino", "minion", "cambria", "bookman")):
        return "Times New Roman", bold, italic, "tibo" if bold else "tiro", False
    return "Helvetica Neue", bold, italic, "hebo" if bold else "helv", False


def _scale_for(lines):
    """Factor so a substituted font occupies about the same width as the PDF's."""
    ratios = []
    for line in lines:
        want = line["bbox"][2] - line["bbox"][0]
        have = 0
        for s in line["spans"]:
            _, _, _, measure, exact = _font_info(s["font"], s["flags"])
            if exact or measure is None:
                return 1.0
            have += pymupdf.Font(measure).text_length(s["text"], s["size"])
        if have > 0 and want > 0:
            ratios.append(want / have)
    if not ratios:
        return 1.0
    return min(max(statistics.median(ratios), 0.8), 1.1)


# --------------------------------------------------------------------------- #
# Images
# --------------------------------------------------------------------------- #
def _image_data(doc, xref):
    """Original bytes when the image is plain RGB/gray JPEG/PNG, else re-encode."""
    try:
        info = doc.extract_image(xref)
    except Exception:
        info = None
    if info and not info.get("smask") and info["ext"] in ("jpeg", "jpg", "png") \
            and info.get("colorspace") in (1, 3):
        return info["image"]

    pix = pymupdf.Pixmap(doc, xref)
    if pix.colorspace and pix.colorspace.n >= 4:  # CMYK -> RGB
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    smask = doc.xref_get_key(xref, "SMask")
    if smask[0] == "xref":
        mask = pymupdf.Pixmap(doc, int(smask[1].split()[0]))
        if (mask.width, mask.height) == (pix.width, pix.height):
            if pix.alpha:
                pix = pymupdf.Pixmap(pix, 0)
            pix = pymupdf.Pixmap(pix, mask)
    return pix.tobytes("png")


def _add_images(slide, page, fit):
    doc = page.parent
    for info in page.get_image_info(xrefs=True):
        xref = info.get("xref")
        if not xref:
            continue
        try:
            data = _image_data(doc, xref)
        except Exception as e:
            print(f"  skipped image {xref}: {e}", file=sys.stderr)
            continue
        r = pymupdf.Rect(info["bbox"]).normalize()
        slide.shapes.add_picture(BytesIO(data), fit.x(r.x0), fit.y(r.y0), fit.d(r.width), fit.d(r.height))


# --------------------------------------------------------------------------- #
# Line art + background
# --------------------------------------------------------------------------- #
def _add_line_art(slide, page, fit):
    """Everything that isn't text or an image, as one transparent picture."""
    drawings = page.get_drawings()
    if not drawings:
        return
    copy = pymupdf.open()
    copy.insert_pdf(page.parent, from_page=page.number, to_page=page.number)
    cp = copy[0]

    # Full-page fills (the page "paper") would make the overlay opaque and hide
    # the pictures below it. Use the topmost one as the slide background and
    # strip them from the overlay; they all touch the corner.
    page_area = page.rect.width * page.rect.height
    full = [d["fill"] for d in drawings if d.get("fill") and d["rect"].get_area() > 0.95 * page_area]
    if full:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(*(round(c * 255) for c in full[-1]))
        cp.add_redact_annot(pymupdf.Rect(0, 0, 1, 1), fill=False)
        cp.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                            graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED,
                            text=pymupdf.PDF_REDACT_TEXT_NONE)
        if not cp.get_drawings():
            copy.close()
            return

    # Drop text and images (they are re-added as real objects), keep line art.
    cp.add_redact_annot(cp.rect, fill=False)
    cp.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_REMOVE,
                        graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
    dpi = min(300, 6000 / (page.rect.width / 72))
    data = cp.get_pixmap(dpi=dpi, alpha=True).tobytes("png")
    copy.close()
    slide.shapes.add_picture(BytesIO(data), fit.x(0), fit.y(0), fit.d(page.rect.width), fit.d(page.rect.height))


# --------------------------------------------------------------------------- #
# Text
# --------------------------------------------------------------------------- #
def _is_rotated(line):
    dx, dy = line["dir"]
    return abs(dy) > 0.02 or dx < 0


def _alignment(lines):
    if len(lines) < 2:
        return PP_ALIGN.LEFT
    x0 = [l["bbox"][0] for l in lines]
    x1 = [l["bbox"][2] for l in lines]
    mid = [(a + b) / 2 for a, b in zip(x0, x1)]
    if max(x0) - min(x0) < 2:
        return PP_ALIGN.LEFT
    if max(x1) - min(x1) < 2:
        return PP_ALIGN.RIGHT
    if max(mid) - min(mid) < 2:
        return PP_ALIGN.CENTER
    return PP_ALIGN.LEFT


def _split_lines(block):
    """PyMuPDF sometimes puts unrelated text (a row of labels, a callout label)
    into one block. Start a new group when a line sits beside, or far off to the
    side of, the group so far; one-column paragraphs stay together. Rotated
    lines each get their own group.
    """
    lines = [l for l in block["lines"] if l["spans"] and "".join(s["text"] for s in l["spans"]).strip()]
    flat = [l for l in lines if not _is_rotated(l)]
    groups = [[l] for l in lines if _is_rotated(l)]
    if _alignment(flat) != PP_ALIGN.LEFT:
        return groups + ([flat] if flat else [])
    cur = None
    for line in flat:
        size = line["spans"][0]["size"]
        if cur:
            same_column = abs(line["bbox"][0] - cur[0]["bbox"][0]) <= 3 * size
            below = line["spans"][0]["origin"][1] > cur[-1]["spans"][0]["origin"][1] + 0.3 * size
            if same_column and below:
                cur.append(line)
                continue
        cur = [line]
        groups.append(cur)
    return groups


def _add_text(slide, page, fit):
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for lines in _split_lines(block):
            _add_text_box(slide, lines, fit)


def _add_text_box(slide, lines, fit):
    x0 = min(l["bbox"][0] for l in lines)
    y0 = min(l["bbox"][1] for l in lines)
    x1 = max(l["bbox"][2] for l in lines)
    y1 = max(l["bbox"][3] for l in lines)
    scale = _scale_for(lines)
    align = _alignment(lines)
    rotation = 0
    left, top, width, height = x0, y0, x1 - x0 + 4, y1 - y0

    if _is_rotated(lines[0]):
        dx, dy = lines[0]["dir"]
        rotation = math.degrees(math.atan2(dy, dx)) % 360
        c, s = abs(dx), abs(dy)
        thick = lines[0]["spans"][0]["size"] * 1.25
        length = max((x1 - x0) * c + (y1 - y0) * s - 2 * thick * s * c, thick)
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        width, height = length + 4, thick
        left, top = cx - width / 2, cy - height / 2

    box = slide.shapes.add_textbox(fit.x(left), fit.y(top), fit.d(width), fit.d(height))
    box.rotation = rotation
    tf = box.text_frame
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    origins = [l["spans"][0]["origin"][1] for l in lines]
    gaps = [b - a for a, b in zip(origins, origins[1:]) if b - a > 0]
    pitch = statistics.median(gaps) if gaps else None

    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        if pitch:
            para.line_spacing = Pt(pitch * fit.s)
        para.space_before = para.space_after = Pt(0)
        for span in line["spans"]:
            if not span["text"]:
                continue
            family, bold, italic, _, exact = _font_info(span["font"], span["flags"])
            run = para.add_run()
            run.text = span["text"].rstrip("\n")
            font = run.font
            font.size = Pt(max(span["size"] * (1 if exact else scale) * fit.s, 1))
            font.name = family
            font.bold = bold
            font.italic = italic
            c = span["color"]
            font.color.rgb = RGBColor((c >> 16) & 255, (c >> 8) & 255, c & 255)


# --------------------------------------------------------------------------- #
# Build + Keynote
# --------------------------------------------------------------------------- #
def pdf_to_pptx(pdf_path, pptx_path, background=True, progress=None):
    doc = pymupdf.open(pdf_path)
    if doc.needs_pass:
        raise ValueError("PDF is password protected")
    if not len(doc):
        raise ValueError("PDF has no pages")
    prs = Presentation()
    first = doc[0].rect
    prs.slide_width, prs.slide_height = _emu(first.width), _emu(first.height)
    blank = prs.slide_layouts[6]

    for i, page in enumerate(doc):
        if progress:
            progress(f"Reading page {i + 1} of {len(doc)}", 0.7 * i / len(doc))
        fit = _Fit(page.rect, first.width, first.height)
        slide = prs.slides.add_slide(blank)
        _add_images(slide, page, fit)
        if background:
            _add_line_art(slide, page, fit)
        _add_text(slide, page, fit)
    prs.save(pptx_path)


def _wait_for_keynote(timeout=30):
    subprocess.run(["open", "-a", "Keynote"], check=True)  # `launch` via AppleScript can fail
    end = time.time() + timeout
    while time.time() < end:
        if subprocess.run(["pgrep", "-x", "Keynote"], capture_output=True).returncode == 0:
            time.sleep(2)
            return
        time.sleep(0.5)
    raise RuntimeError("Keynote did not start")


def pptx_to_keynote(pptx_path, key_path, png_dir=None):
    """Open the .pptx in Keynote, save it as .key and optionally export slide PNGs."""
    script = """
    on run argv
        set inFile to POSIX file (item 1 of argv)
        set outFile to POSIX file (item 2 of argv)
        tell application "Keynote"
            set d to open inFile
            save d in outFile
            if (count of argv) > 2 then
                export d to POSIX file (item 3 of argv) as slide images
            end if
            close d saving no
        end tell
    end run
    """
    _wait_for_keynote()
    args = ["osascript", "-e", script, str(pptx_path), str(key_path)]
    if png_dir:
        args.append(str(png_dir))
    r = subprocess.run(args, capture_output=True, text=True, timeout=900)
    if r.returncode:
        raise RuntimeError(f"Keynote conversion failed: {r.stderr.strip()}")


# --------------------------------------------------------------------------- #
# Accuracy check
# --------------------------------------------------------------------------- #
@dataclass
class SlideReport:
    page: int
    match: float       # 0-100, higher is better
    off_pixels: float  # % of pixels that differ noticeably


def compare(pdf_path, png_dir, out_dir, thumb_width=720):
    """Compare each exported Keynote slide with the PDF page; save thumbnails."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    pngs = sorted(Path(png_dir).glob("*.png"))
    doc = pymupdf.open(pdf_path)
    reports = []
    for i, page in enumerate(doc):
        if i >= len(pngs):
            break
        key_img = Image.open(pngs[i]).convert("RGB")
        w, h = key_img.size
        pix = page.get_pixmap(matrix=pymupdf.Matrix(w / page.rect.width, h / page.rect.height), alpha=False)
        pdf_img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        if pdf_img.size != key_img.size:
            pdf_img = pdf_img.resize(key_img.size)
        diff = np.abs(np.asarray(pdf_img, dtype=np.int16) - np.asarray(key_img, dtype=np.int16))
        match = 100 * (1 - diff.mean() / 255)
        off = 100 * (diff.max(axis=2) > 48).mean()
        reports.append(SlideReport(i + 1, round(match, 2), round(off, 2)))

        heat = Image.fromarray(np.clip(diff.max(axis=2) * 3, 0, 255).astype(np.uint8)).convert("RGB")
        th = round(thumb_width * h / w)
        for name, im in (("pdf", pdf_img), ("key", key_img), ("diff", heat)):
            im.resize((thumb_width, th)).save(out_dir / f"{i + 1}_{name}.jpg", quality=85)
    return reports


def missing_fonts(pdf_path):
    """Font families used by the PDF that aren't installed (they get substituted)."""
    missing = set()
    doc = pymupdf.open(pdf_path)
    for page in doc:
        for f in page.get_fonts():
            family, _, _, _, exact = _font_info(f[3], 0)
            if not exact:
                base = re.split(r"[-,]", f[3].split("+")[-1])[0]
                missing.add(re.sub(r"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", " ", base))
    return sorted(missing)


# --------------------------------------------------------------------------- #
# Public entry point
# --------------------------------------------------------------------------- #
def convert(pdf_path, key_path, background=True, verify=True, progress=None, preview_dir=None):
    """PDF -> .key. Returns a list of SlideReport (empty if verify is off)."""
    pdf_path, key_path = Path(pdf_path), Path(key_path)
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pptx = tmp / "deck.pptx"
        pdf_to_pptx(pdf_path, pptx, background, progress)
        if progress:
            progress("Importing into Keynote", 0.8)
        png_dir = tmp / "slides" if verify else None
        pptx_to_keynote(pptx, key_path, png_dir)
        if not verify:
            return []
        if progress:
            progress("Comparing slides with the PDF", 0.95)
        return compare(pdf_path, png_dir, preview_dir or tmp / "preview")
