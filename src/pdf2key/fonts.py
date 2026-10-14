"""Identify fonts by their glyph widths.

PDFs from browsers often embed fonts without names (Type3 glyph procedures).
Every font face has its own set of advance widths, so measuring the distance
between consecutive glyphs in the PDF and comparing it with each installed
face picks out the family, weight and italic. Variable fonts are measured at
each of their named instances.

The catalog of installed faces is cached in ~/Library/Caches/pdf2key.
"""
import json
import re
from dataclasses import dataclass
from pathlib import Path

import numpy as np

FONT_DIRS = [Path("/Library/Fonts"), Path("/System/Library/Fonts"), Path.home() / "Library/Fonts"]
CACHE = Path.home() / "Library/Caches/pdf2key/fonts-v3.json"
# Latin, punctuation, currency, arrows: enough to fingerprint text faces
CODEPOINTS = [*range(0x20, 0x250), *range(0x2000, 0x20D0), *range(0x2190, 0x2200)]


@dataclass
class Face:
    family: str     # typographic family, e.g. "Inter"
    style: str      # e.g. "SemiBold Italic"
    weight: int     # 100-900
    italic: bool
    mono: bool
    variable: bool  # an instance of a variable font
    system: bool    # ships with macOS (vs. installed by the user)
    path: str       # font file, and its index inside a .ttc collection
    index: int
    adv: dict       # codepoint -> advance in em

    @property
    def keynote_name(self):
        """Name Keynote resolves to this exact face."""
        style = re.sub(r"\b(Regular|Italic|Roman|Book)\b", "", self.style).strip()
        return f"{self.family} {style}".strip()



def _faces_in(tt, path, index=0):
    name = tt["name"]
    family = name.getDebugName(16) or name.getDebugName(1)
    cmap = tt.getBestCmap() or {}
    upm = tt["head"].unitsPerEm
    mono = bool(tt["post"].isFixedPitch) if "post" in tt else False
    system = str(path).startswith("/System/")
    wanted = {cp: g for cp, g in cmap.items() if cp in _CP_SET}
    if "fvar" in tt:
        for inst in tt["fvar"].instances:
            style = name.getDebugName(inst.subfamilyNameID) or "Regular"
            gs = tt.getGlyphSet(location=inst.coordinates, normalized=False)
            adv = {cp: round(gs[g].width / upm, 5) for cp, g in wanted.items() if g in gs}
            italic = "italic" in style.lower() or "italic" in path.name.lower()
            yield Face(family, style, int(inst.coordinates.get("wght", 400)), italic, mono, True, system,
                       str(path), index, adv)
    else:
        style = name.getDebugName(17) or name.getDebugName(2) or "Regular"
        hm = tt["hmtx"].metrics
        adv = {cp: round(hm[g][0] / upm, 5) for cp, g in wanted.items() if g in hm}
        weight = tt["OS/2"].usWeightClass if "OS/2" in tt else 400
        italic = bool(tt["head"].macStyle & 2) or "italic" in style.lower()
        yield Face(family, style, weight, italic, mono, False, system, str(path), index, adv)


_CP_SET = set(CODEPOINTS)
_catalog = None


def catalog():
    """All installed text faces (cached on disk, refreshed when files change)."""
    global _catalog
    if _catalog is not None:
        return _catalog
    import logging

    from fontTools.ttLib import TTCollection, TTFont
    logging.getLogger("fontTools").setLevel(logging.ERROR)  # noisy about quirky system fonts

    try:
        cache = json.loads(CACHE.read_text())
    except Exception:
        cache = {}
    fresh, faces = {}, []
    for d in FONT_DIRS:
        if not d.exists():
            continue
        for f in d.rglob("*"):
            if f.suffix.lower() not in (".ttf", ".otf", ".ttc"):
                continue
            key = f"{f}|{f.stat().st_mtime_ns}|{f.stat().st_size}"
            if key in cache:
                entries = cache[key]
            else:
                entries = []
                try:
                    fonts = TTCollection(f, lazy=True).fonts if f.suffix.lower() == ".ttc" else [TTFont(f, lazy=True)]
                    for i, tt in enumerate(fonts):
                        entries += [face.__dict__ for face in _faces_in(tt, f, i)]
                except Exception:
                    pass  # unreadable or bitmap-only font
            fresh[key] = entries
            faces += [Face(**{**e, "adv": {int(k): v for k, v in e["adv"].items()}}) for e in entries]
    if fresh.keys() != cache.keys():
        try:
            CACHE.parent.mkdir(parents=True, exist_ok=True)
            CACHE.write_text(json.dumps(fresh))
        except Exception:
            pass
    # skip symbol/pi fonts and macOS's hidden UI fonts (".LastResort", ".SF …")
    _catalog = [f for f in faces if len(f.adv) > 60 and not f.family.startswith(".")]
    return _catalog


@dataclass
class Match:
    face: Face
    error: float     # mean |measured - predicted| advance, in em
    tracking: float  # letter-spacing, in em


def identify(samples, slanted=None, faces=None):
    """samples: (codepoint, size, measured advance, run id). Letter-spacing is
    fitted per run, since one font is often used at several spacings.
    slanted: whether the glyphs looked italic (breaks ties between an upright
    and italic face with the same widths).

    Returns the candidates that fit equally well, best guess first; empty if
    no installed face fits."""
    faces = faces or catalog()
    if len(samples) < 6:
        return []
    cps = np.array([s[0] for s in samples])
    size = np.array([s[1] for s in samples], dtype=float)
    act = np.array([s[2] for s in samples], dtype=float) / size
    runs = np.array([s[3] for s in samples])
    groups = [runs == g for g in np.unique(runs)]
    keep = max(6, int(len(samples) * 0.8))  # ignore the worst 20%: word gaps, kerning
    scored = []
    for face in faces:
        try:
            pred = np.array([face.adv[cp] for cp in cps])
        except KeyError:
            continue
        r = act - pred
        t = np.empty_like(r)
        for g in groups:
            t[g] = np.median(r[g])
        err = float(np.mean(np.sort(np.abs(r - t))[:keep]))
        scored.append(Match(face, err, float(np.median(t))))
    if not scored:
        return []
    best = min(m.error for m in scored)
    if best > 0.01:
        return []
    # Faces whose widths fit equally well (all monospace fonts; an italic and
    # its upright; a family's static and variable copies): prefer the italic
    # the glyphs suggest, static files, and fonts the user installed.
    tied = [m for m in scored if m.error <= best + 0.002]
    return sorted(tied, key=lambda m: (m.error > best + 0.0005 and not m.face.mono,
                                       slanted is not None and m.face.italic != slanted,
                                       m.face.variable, m.face.system, m.error))
