"""Experimental: HTML deck -> Keynote with movable pictures as well as editable text.

A PDF only says "draw this image here"; the HTML a deck was made from says
which element is a photo, how it is cropped, and what is layered over it. So:

1. Chrome (headless, via Playwright) lays out each slide on its own page at
   the deck's size, with the deck's fonts.
2. Every element is listed in paint order and sorted into
   - pictures: <img> and CSS background images. They become real Keynote
     pictures from the original files, with the crop the browser applies
     (object-fit / object-position / background-size, overflow clipping),
     opacity and rounded corners. Pictures with effects Keynote can't
     reproduce (filters, masks, rotation, blend modes) are captured on
     their own instead, still as a separate, movable picture.
   - decorations: backgrounds, cards, gradients, borders, SVG. Each run of
     decorations between two pictures is captured as one transparent layer,
     trimmed to its content, so overlays stay above the photos they cover.
3. Chrome also prints each slide to PDF, and the regular text pipeline turns
   that into editable text boxes on top.

Slides are the deck's <section> elements (Claude Design exports, reveal-style
decks); a page without sections becomes a single slide.
"""
import base64
import json
import re
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path
from urllib.parse import unquote, urlparse

import pymupdf
from PIL import Image

from pptx.oxml.ns import qn

from . import converter

PX_TO_PT = 0.75
SCALE = 2           # screenshot pixels per CSS pixel
MAX_IMAGE_SIDE = 3200

# Everything the page knows about each element, in paint order.
COLLECT_JS = r"""
(root) => {
  const items = [];
  const all = [root, ...root.querySelectorAll('*')];
  const rootRect = root.getBoundingClientRect();
  const R = r => ({x: r.left - rootRect.left, y: r.top - rootRect.top, w: r.width, h: r.height});
  let order = 0;
  for (const el of all) {
    el.setAttribute('data-p2k', String(order++));
  }
  const inSvg = el => el.closest('svg') && el.tagName.toLowerCase() !== 'svg';
  const effZ = el => {
    for (let e = el; e && e !== root.parentElement; e = e.parentElement) {
      const cs = getComputedStyle(e);
      if (cs.position !== 'static' && cs.zIndex !== 'auto') return parseInt(cs.zIndex) || 0;
    }
    return 0;
  };
  const complex = el => {   // effects a Keynote picture can't reproduce
    for (let e = el; e && e !== root.parentElement; e = e.parentElement) {
      const cs = getComputedStyle(e);
      if (cs.filter !== 'none' || cs.mixBlendMode !== 'normal' || cs.clipPath !== 'none'
          || (cs.maskImage && cs.maskImage !== 'none') || (cs.webkitMaskImage && cs.webkitMaskImage !== 'none'))
        return true;
      if (cs.transform !== 'none') {
        const m = new DOMMatrix(cs.transform);
        if (Math.abs(m.b) > 1e-3 || Math.abs(m.c) > 1e-3) return true;   // rotated / skewed
      }
    }
    return false;
  };
  const opacity = el => {
    let o = 1;
    for (let e = el; e && e !== root.parentElement; e = e.parentElement) o *= parseFloat(getComputedStyle(e).opacity);
    return o;
  };
  const clipBox = el => {   // visible area after overflow clipping by ancestors
    let c = {l: -1e9, t: -1e9, r: 1e9, b: 1e9}, radius = null;
    for (let e = el.parentElement; e && e !== root.parentElement; e = e.parentElement) {
      const cs = getComputedStyle(e);
      if (cs.overflow !== 'visible' || cs.overflowX !== 'visible' || cs.overflowY !== 'visible') {
        const r = e.getBoundingClientRect();
        if (r.width * r.height < (c.r - c.l) * (c.b - c.t)) radius = parseFloat(cs.borderTopLeftRadius) || 0;
        c = {l: Math.max(c.l, r.left), t: Math.max(c.t, r.top), r: Math.min(c.r, r.right), b: Math.min(c.b, r.bottom)};
      }
    }
    return {x: c.l - rootRect.left, y: c.t - rootRect.top, w: c.r - c.l, h: c.b - c.t, radius};
  };
  for (const el of all) {
    if (inSvg(el)) continue;
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden') continue;
    const r = el.getBoundingClientRect();
    if (r.width < 0.5 || r.height < 0.5) continue;
    const tag = el.tagName.toLowerCase();
    const base = {id: el.getAttribute('data-p2k'), tag, rect: R(r), z: effZ(el), order: +el.getAttribute('data-p2k')};
    const bg = cs.backgroundImage;
    const urls = bg && bg !== 'none' ? [...bg.matchAll(/url\("?([^")]*)"?\)/g)].map(m => m[1]) : [];
    if (tag === 'img' && el.currentSrc && el.naturalWidth) {
      items.push({...base, kind: 'image', src: el.currentSrc, natural: [el.naturalWidth, el.naturalHeight],
        fit: cs.objectFit, pos: cs.objectPosition, opacity: opacity(el), complex: complex(el),
        radius: parseFloat(cs.borderTopLeftRadius) || 0, clip: clipBox(el),
        content: R(new DOMRect(r.left + parseFloat(cs.borderLeftWidth) + parseFloat(cs.paddingLeft),
                               r.top + parseFloat(cs.borderTopWidth) + parseFloat(cs.paddingTop),
                               el.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight),
                               el.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom)))});
      // the img's own box (border, background) still paints as decoration
    }
    if (urls.length === 1 && !/gradient\(/.test(bg)) {
      items.push({...base, kind: 'bgimage', src: urls[0], size: cs.backgroundSize, pos: cs.backgroundPosition,
        repeat: cs.backgroundRepeat, opacity: opacity(el), complex: complex(el),
        radius: parseFloat(cs.borderTopLeftRadius) || 0, clip: clipBox(el)});
    }
    const pseudo = ['::before', '::after'].some(p => getComputedStyle(el, p).content !== 'none'
                                                   && getComputedStyle(el, p).content !== 'normal');
    const paints = tag !== 'img' && (el === root || tag === 'svg' || tag === 'canvas' || tag === 'video' || tag === 'hr'
      || tag === 'iframe' || pseudo
      || !/rgba\(0, 0, 0, 0\)|transparent/.test(cs.backgroundColor)
      || (bg !== 'none' && (urls.length !== 1 || /gradient\(/.test(bg)))
      || ['Top', 'Right', 'Bottom', 'Left'].some(s => parseFloat(cs['border' + s + 'Width']) > 0
                                                && cs['border' + s + 'Style'] !== 'none')
      || cs.boxShadow !== 'none' || (cs.outlineStyle !== 'none' && parseFloat(cs.outlineWidth) > 0));
    if (paints) items.push({...base, kind: 'decoration'});
  }
  items.sort((a, b) => a.z - b.z || a.order - b.order);
  return items;
}
"""

BASE_CSS = """
*, *::before, *::after { animation: none !important; transition: none !important; }
html, body { margin: 0 !important; padding: 0 !important; background: transparent !important; overflow: hidden !important; }
"""
HIDE_TEXT_CSS = """
[data-p2k-root], [data-p2k-root] * { -webkit-text-fill-color: transparent !important; text-shadow: none !important;
  -webkit-text-stroke: 0 !important; caret-color: transparent !important; text-decoration-color: transparent !important; }
[data-p2k-root] svg text { fill: transparent !important; stroke: none !important; }
"""


def _find_html(path, workdir):
    path = Path(path)
    if path.suffix.lower() == ".zip":
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():  # no absolute paths or ".." escapes
                if name.startswith("/") or ".." in Path(name).parts:
                    raise ValueError(f"unsafe path in zip: {name}")
            z.extractall(workdir)
        candidates = [p for p in Path(workdir).rglob("*.html") if not p.name.startswith(".")]
        if not candidates:
            raise ValueError("no .html file in the zip")
        path = max(candidates, key=lambda p: p.stat().st_size)
    return path.resolve()


def _slide_size(page):
    size = page.evaluate("""() => {
      const x = document.querySelector('[width][height]:not(img):not(svg):not(canvas):not(video):not(iframe)');
      return x ? [parseInt(x.getAttribute('width')), parseInt(x.getAttribute('height'))] : null }""")
    return tuple(size) if size and all(size) else (1920, 1080)


def _head_markup(page):
    """Stylesheets and fonts of the deck, to rebuild each slide on its own page."""
    return page.evaluate("""() => [...document.querySelectorAll('link[rel=stylesheet], style')]
                              .map(e => e.outerHTML).join('\\n')""")


def _render_layer(page, visible, hide_bg_images):
    """Screenshot with only `visible` elements painted (no text, no pictures)."""
    rules = ["[data-p2k-root], [data-p2k-root] *:not(svg *) { visibility: hidden !important; }"]
    if visible:
        # :not(#p2k) adds an id's worth of specificity, so these beat the rule above
        sel = ", ".join(f'[data-p2k="{i}"]:not(#p2k)' for i in visible)
        rules.append(f"{sel} {{ visibility: visible !important; }}")
        rules.append(", ".join(f'[data-p2k="{i}"] svg *:not(#p2k)' for i in visible)
                     + " { visibility: visible !important; }")
    if hide_bg_images:
        rules.append(", ".join(f'[data-p2k="{i}"]:not(#p2k)' for i in hide_bg_images)
                     + " { background-image: none !important; }")
    page.evaluate("css => { document.getElementById('p2k-layer').textContent = css }", "\n".join(rules))
    png = page.screenshot(omit_background=True, type="png")
    return Image.open(BytesIO(png)).convert("RGBA")


def _local_file(src, base_dir):
    u = urlparse(src)
    if u.scheme == "file":
        return Path(unquote(u.path))
    if u.scheme in ("", None):
        return (base_dir / unquote(src)).resolve()
    if u.scheme == "data":
        return None
    return None


def _image_bytes(src, base_dir):
    if src.startswith("data:"):
        return base64.b64decode(src.split(",", 1)[1]) if ";base64," in src[:100] else None
    f = _local_file(src, base_dir)
    return f.read_bytes() if f and f.exists() else None


def _parse_pos(value, free):
    """CSS position component ('50%', '12px', 'left') -> offset in px."""
    value = value.strip()
    keywords = {"left": 0, "top": 0, "center": 50, "right": 100, "bottom": 100}
    if value in keywords:
        return free * keywords[value] / 100
    if value.endswith("%"):
        return free * float(value[:-1]) / 100
    if value.endswith("px"):
        return float(value[:-2])
    return free / 2


def _displayed_rect(item):
    """Where the whole image is drawn (x, y, w, h), before clipping."""
    iw, ih = item["natural"]
    if item["kind"] == "image":
        box = item["content"]
        fit = item["fit"]
        pos = item["pos"].split()
    else:
        box = item["rect"]
        size = item["size"]
        fit = "cover" if size == "cover" else "contain" if size == "contain" else "explicit"
        pos = item["pos"].split()
        if fit == "explicit":
            parts = size.split()
            def dim(v, total, other_auto_ratio=None):
                if v.endswith("%"):
                    return total * float(v[:-1]) / 100
                if v.endswith("px"):
                    return float(v[:-2])
                return None
            w = dim(parts[0], box["w"]) if parts else None
            h = dim(parts[1], box["h"]) if len(parts) > 1 else None
            if w is None and h is None:
                w, h = iw, ih
            elif w is None:
                w = h * iw / ih
            elif h is None:
                h = w * ih / iw
            px, py = (pos + ["50%", "50%"])[:2]
            return (box["x"] + _parse_pos(px, box["w"] - w), box["y"] + _parse_pos(py, box["h"] - h), w, h)
    bw, bh = box["w"], box["h"]
    if fit == "cover":
        k = max(bw / iw, bh / ih)
    elif fit == "contain":
        k = min(bw / iw, bh / ih)
    elif fit == "none":
        k = 1
    elif fit == "scale-down":
        k = min(1, bw / iw, bh / ih)
    else:  # fill
        return (box["x"], box["y"], bw, bh)
    w, h = iw * k, ih * k
    px, py = (pos + ["50%", "50%"])[:2]
    return (box["x"] + _parse_pos(px, bw - w), box["y"] + _parse_pos(py, bh - h), w, h)


def _add_picture(slide, item, base_dir, layer_png=None):
    """A real picture: original file, cropped like the browser crops it."""
    k = PX_TO_PT
    if layer_png is not None:  # captured on its own (effects Keynote can't do)
        bbox = layer_png.getbbox()
        if not bbox:
            return
        crop = layer_png.crop(bbox)
        buf = BytesIO()
        crop.save(buf, "PNG")
        x0, y0 = bbox[0] / SCALE, bbox[1] / SCALE
        slide.shapes.add_picture(BytesIO(buf.getvalue()), converter._emu(x0 * k), converter._emu(y0 * k),
                                 converter._emu(crop.width / SCALE * k), converter._emu(crop.height / SCALE * k))
        return
    data = _image_bytes(item["src"], base_dir)
    if not data:
        return
    img = Image.open(BytesIO(data))
    item.setdefault("natural", list(img.size))
    dx, dy, dw, dh = _displayed_rect(item)
    # visible part: displayed image ∩ its box ∩ ancestors' overflow clip
    box = item["content"] if item["kind"] == "image" else item["rect"]
    c = item["clip"]
    vx0, vy0 = max(dx, box["x"], c["x"]), max(dy, box["y"], c["y"])
    vx1 = min(dx + dw, box["x"] + box["w"], c["x"] + c["w"])
    vy1 = min(dy + dh, box["y"] + box["h"], c["y"] + c["h"])
    if vx1 - vx0 < 0.5 or vy1 - vy0 < 0.5:
        return

    # shrink huge originals to what the slide can show at retina density
    need = max((vx1 - vx0) / dw * img.width, (vy1 - vy0) / dh * img.height)
    show_px = max(vx1 - vx0, vy1 - vy0) * SCALE
    if max(img.size) > MAX_IMAGE_SIDE and need > show_px * 1.5:
        f = max(show_px * 1.5 / need, MAX_IMAGE_SIDE / max(img.size))
        img = img.resize((max(1, round(img.width * f)), max(1, round(img.height * f))), Image.LANCZOS)
        buf = BytesIO()
        if img.mode in ("RGBA", "LA", "P"):
            img.save(buf, "PNG", optimize=True)
        else:
            img.convert("RGB").save(buf, "JPEG", quality=90)
        data = buf.getvalue()

    pic = slide.shapes.add_picture(BytesIO(data), converter._emu(vx0 * k), converter._emu(vy0 * k),
                                   converter._emu((vx1 - vx0) * k), converter._emu((vy1 - vy0) * k))
    pic.crop_left = (vx0 - dx) / dw
    pic.crop_top = (vy0 - dy) / dh
    pic.crop_right = (dx + dw - vx1) / dw
    pic.crop_bottom = (dy + dh - vy1) / dh

    if item["opacity"] < 0.999:
        blip = pic._element.blipFill.find(qn("a:blip"))
        blip.append(blip.makeelement(qn("a:alphaModFix"), {"amt": str(round(item["opacity"] * 100000))}))
    radius = item["radius"] or (c.get("radius") or 0)
    short = min(vx1 - vx0, vy1 - vy0)
    if radius > 0.5 and short > 0:
        # rounded-rectangle mask; adj is the radius as a share of the short side (max 50000)
        geom = pic._element.spPr.find(qn("a:prstGeom"))
        geom.set("prst", "roundRect")
        av = geom.find(qn("a:avLst"))
        if av is None:
            av = geom.makeelement(qn("a:avLst"), {})
            geom.append(av)
        av.append(av.makeelement(qn("a:gd"), {"name": "adj", "fmla": f"val {round(min(0.5, radius / short) * 100000)}"}))


def _add_layer(slide, png):
    bbox = png.getbbox()
    if not bbox:
        return
    crop = png.crop(bbox)
    buf = BytesIO()
    crop.save(buf, "PNG", optimize=True)
    k = PX_TO_PT / SCALE
    slide.shapes.add_picture(BytesIO(buf.getvalue()), converter._emu(bbox[0] * k), converter._emu(bbox[1] * k),
                             converter._emu(crop.width * k), converter._emu(crop.height * k))


def html_to_pptx(html_path, pptx_path, pdf_out, progress=None, font_report=None):
    """Build the .pptx from an HTML deck (or a zip of one). Also writes the
    deck as a PDF to `pdf_out` (used for the text and the accuracy check)."""
    from playwright.sync_api import sync_playwright

    with tempfile.TemporaryDirectory() as tmp, sync_playwright() as pw:
        html = _find_html(html_path, tmp)
        base_dir = html.parent
        browser = pw.chromium.launch(channel="chrome")
        try:
            page = browser.new_page(viewport={"width": 1920, "height": 1080})
            page.goto(html.as_uri(), wait_until="load")
            page.wait_for_timeout(1500)
            width, height = _slide_size(page)
            head = _head_markup(page)
            sections = page.evaluate("""() => {
              const s = [...document.querySelectorAll('section')].filter(e => !e.parentElement.closest('section'));
              return s.length ? s.map(e => e.outerHTML) : [document.body.innerHTML] }""")

            slide_page = browser.new_page(viewport={"width": width, "height": height}, device_scale_factor=SCALE)
            slide_page.emulate_media(media="screen")
            pdf = pymupdf.open()
            plans = []
            slide_file = base_dir / f".pdf2key-slide-{next(tempfile._get_candidate_names())}.html"
            for i, markup in enumerate(sections):
                if progress:
                    progress(f"Laying out slide {i + 1} of {len(sections)}", 0.25 * i / len(sections))
                doc_html = f"""<!doctype html><html><head><meta charset="utf-8">
                    {head}<style>{BASE_CSS}</style><style id="p2k-layer"></style></head><body>{markup}</body></html>"""
                # a file next to the deck, so relative images and fonts load (file:// pages
                # can't be reached from about:blank)
                slide_file.write_text(doc_html, encoding="utf-8")
                slide_page.goto(slide_file.as_uri(), wait_until="load")
                slide_page.evaluate(f"""() => {{
                    const s = document.body.firstElementChild;
                    s.setAttribute('data-p2k-root', '');
                    Object.assign(s.style, {{position: 'absolute', left: '0', top: '0', width: '{width}px',
                        height: '{height}px', opacity: '1', transform: 'none', visibility: 'visible',
                        display: getComputedStyle(s).display === 'none' ? 'block' : getComputedStyle(s).display}});
                }}""")
                slide_page.evaluate("() => document.fonts.ready")
                slide_page.wait_for_timeout(150)
                pdf_bytes = slide_page.pdf(width=f"{width}px", height=f"{height}px", print_background=True,
                                           page_ranges="1", margin={"top": "0", "right": "0", "bottom": "0", "left": "0"})
                pdf.insert_pdf(pymupdf.open("pdf", pdf_bytes))

                items = slide_page.evaluate(COLLECT_JS, slide_page.query_selector("[data-p2k-root]"))
                slide_page.add_style_tag(content=HIDE_TEXT_CSS)
                bg_image_ids = [it["id"] for it in items if it["kind"] == "bgimage"]
                layers, run = [], []
                for it in items:
                    if it["kind"] == "decoration":
                        run.append(it["id"])
                        continue
                    if run:
                        layers.append(("layer", _render_layer(slide_page, run, bg_image_ids)))
                        run = []
                    if it["complex"]:
                        layers.append(("picture", it, _render_layer(slide_page, [it["id"]], [])))
                    else:
                        layers.append(("picture", it, None))
                if run:
                    layers.append(("layer", _render_layer(slide_page, run, bg_image_ids)))
                plans.append(layers)
            pdf.save(pdf_out)
        finally:
            browser.close()
            if "slide_file" in locals():
                slide_file.unlink(missing_ok=True)

        doc = converter.open_pdf(pdf_out)
        pages_blocks = converter.analyze_text(doc, None, font_report)
        w_pt, h_pt = width * PX_TO_PT, height * PX_TO_PT
        prs = converter.new_presentation(w_pt, h_pt)
        blank = prs.slide_layouts[6]
        for i, layers in enumerate(plans):
            if progress:
                progress(f"Building slide {i + 1} of {len(plans)}", 0.4 + 0.3 * i / len(plans))
            slide = prs.slides.add_slide(blank)
            for layer in layers:
                if layer[0] == "layer":
                    _add_layer(slide, layer[1])
                else:
                    _add_picture(slide, layer[1], base_dir, layer[2])
            converter.add_text(slide, pages_blocks[i], doc[i].rect, w_pt, h_pt)
        prs.save(pptx_path)


def convert_html(html_path, key_path, verify=True, progress=None, preview_dir=None, font_report=None):
    """HTML deck (or zip) -> .key. Returns SlideReports like converter.convert."""
    key_path = Path(key_path).resolve()
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pptx, pdf = tmp / "deck.pptx", tmp / "deck.pdf"
        html_to_pptx(html_path, pptx, pdf, progress, font_report)
        if progress:
            progress("Importing into Keynote", 0.8)
        png_dir = tmp / "slides" if verify else None
        converter.pptx_to_keynote(pptx, key_path, png_dir)
        if not verify:
            return []
        if progress:
            progress("Comparing slides with the original", 0.95)
        return converter.compare(pdf, png_dir, preview_dir or tmp / "preview")
