"""The export folder: the converted deck plus everything someone on another
computer needs to see it exactly as intended.

    <name> - pdf2key/
      <name>.key, <name>.pptx
      README.txt            what to install on a fresh Mac or PC, and why
      fonts/                the open-licensed fonts the deck uses, ready to install
        FONTS-LICENSE.txt
      results/              the conversion report and per-slide comparisons
"""
import json
import re
import shutil
from pathlib import Path

from . import report as reporting

# Fonts every Windows PC has, and what PowerPoint for Windows shows for common
# Mac-only fonts.
ON_WINDOWS = {"arial", "times new roman", "georgia", "verdana", "courier new", "trebuchet ms", "tahoma",
              "calibri", "cambria", "consolas", "segoe ui", "comic sans ms", "impact"}
WINDOWS_STANDIN = {"helvetica neue": "Arial", "helvetica": "Arial", "menlo": "Consolas", "monaco": "Consolas",
                   "sf pro": "Segoe UI", "avenir": "Segoe UI", "avenir next": "Segoe UI", "futura": "Century Gothic",
                   "gill sans": "Gill Sans MT", "optima": "Candara", "lucida grande": "Segoe UI"}
FONT_FILE_TYPES = ".ttf, .otf"


def folder_for(source, parent=None):
    """Default export folder next to the input: '<name> - pdf2key'."""
    source = Path(source)
    stem = source.name if source.is_dir() else source.stem
    return Path(parent or source.parent) / f"{stem} - pdf2key"


def _slug(text):
    return re.sub(r"[^A-Za-z0-9]+", "", text) or "Font"


def _family_key(u):
    face = u.get("face")
    return face["family"] if face else re.sub(r"\s+(Thin|ExtraLight|Light|Medium|SemiBold|ExtraBold|Black|Heavy)$",
                                              "", u["name"])


def _ship_fonts(usage, fonts_dir):
    """Copy (or cut from variable fonts) every open-licensed face the deck uses.
    Returns {family: {"files": [...], "license": ..., "url": ..., "copyright": ...}}
    for the faces shipped, and the set of families that couldn't be shipped."""
    from fontTools.ttLib import TTFont

    from .fonts import Face, license_of

    shipped, kept_back, done = {}, set(), set()
    for u in usage:
        f = u.get("face")
        if not f:
            continue
        face = Face(**f, adv={})
        family = face.family
        if face.system:
            continue  # ships with macOS; covered in the README, never copied
        text, url, copyright_, free, reserved = license_of(face)
        if not free:
            kept_back.add(family)
            continue
        entry = shipped.setdefault(family, {"files": [], "license": text, "url": url, "copyright": copyright_})
        fonts_dir.mkdir(parents=True, exist_ok=True)
        if face.variable and not reserved:
            # A static file per weight used: Office apps can't pick a named weight
            # out of a variable font, but they find "Montserrat ExtraBold" in a
            # static file whose family name is exactly that.
            name = f"{_slug(family)}-{_slug(face.style)}.ttf"
            if name in done:
                continue
            try:
                from fontTools.varLib import instancer
                static = instancer.instantiateVariableFont(TTFont(face.path), face.coords, updateFontNames=True)
                static.save(fonts_dir / name)
            except Exception:
                name = Path(face.path).name  # fall back to the variable font itself
                if name not in done:
                    shutil.copyfile(face.path, fonts_dir / name)
        else:
            name = Path(face.path).name
            if name not in done:
                shutil.copyfile(face.path, fonts_dir / name)
        done.add(name)
        if name not in entry["files"]:
            entry["files"].append(name)
    return shipped, kept_back


def _license_file(shipped, path):
    lines = ["Fonts in this folder", "====================", "",
             "These fonts are free and open-source. They are included so the presentation",
             "looks right on any computer. Each is redistributed unmodified or as a static",
             "weight cut from its variable font, under its own license, which also travels",
             "inside every font file.", ""]
    for family, e in sorted(shipped.items()):
        lines += [family, "-" * len(family), f"Files: {', '.join(sorted(e['files']))}"]
        if e["copyright"]:
            lines.append(f"Copyright: {e['copyright']}")
        if e["license"]:
            lines.append(f"License: {e['license']}")
        if e["url"]:
            lines.append(f"Full license: {e['url']}")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def _readme(title, report, outputs, usage, shipped, kept_back):
    acc = report.get("accuracy")
    has_key, has_pptx = "key" in outputs, "pptx" in outputs
    need = [u for u in usage if u.get("face") and not u["face"]["system"]]
    families = []
    for u in usage:
        fam = _family_key(u)
        if fam not in families:
            families.append(fam)

    def styles_of(fam):
        return ", ".join(sorted({u["display"].replace(fam, "").strip() or "Regular"
                                 for u in usage if _family_key(u) == fam}))

    def chars_of(fam):
        return sum(u["chars"] for u in usage if _family_key(u) == fam)

    mac_only = [fam for fam in families
                if any(u.get("face") and u["face"]["system"] for u in usage if _family_key(u) == fam)
                and fam.lower() not in ON_WINDOWS]
    missing = [fam for fam in families if not any(u.get("face") for u in usage if _family_key(u) == fam)]

    L = []
    w = L.append
    w(f"{title}")
    w("=" * len(title))
    w("")
    w("This folder holds an editable version of the presentation, converted by pdf2key,")
    w("and what you need to make it look exactly as designed on your computer.")
    w("")
    w("What's in this folder")
    w("---------------------")
    if has_key:
        w(f"  {outputs['key'].name:<34} Apple Keynote presentation (Mac, iPad, iPhone, iCloud)")
    if has_pptx:
        w(f"  {outputs['pptx'].name:<34} PowerPoint presentation (also opens in Keynote and Google Slides)")
    if shipped:
        w(f"  {'fonts/':<34} the fonts this presentation uses; install them first (below)")
    w(f"  {'results/':<34} how the conversion went, slide by slide")
    w(f"  {'README.txt':<34} this file")
    w("")

    step = 1
    w("Before you open it")
    w("------------------")
    if shipped:
        w(f"{step}. Install the fonts in the fonts/ folder.")
        w("   The presentation uses these fonts by name. If they aren't installed, your")
        w("   app quietly swaps in a different font and headings look wrong (wrong")
        w("   weight, wider or narrower, lines wrapping differently).")
        w("")
        w("   On a Mac:     open Font Book (in Applications), then drag every font file")
        w(f"                 ({FONT_FILE_TYPES}) from the fonts/ folder into its window.")
        w("                 (Or double-click one file at a time and click Install.)")
        w("   On Windows:   open the fonts/ folder, select all the font files, right-click,")
        w("                 and choose \"Install for all users\" (or \"Install\").")
        w("   Already have some of these fonts? Install anyway; nothing is lost.")
        w("")
        step += 1
        w(f"{step}. Quit Keynote / PowerPoint completely, then open them again.")
        w("   Apps only notice newly installed fonts when they start.")
        w("")
        step += 1
    w(f"{step}. Open the presentation.")
    if has_key:
        w(f"   Keynote:      double-click {outputs['key'].name}.")
    if has_pptx:
        w(f"   PowerPoint:   double-click {outputs['pptx'].name}.")
    w("")

    w("Fonts this presentation uses")
    w("----------------------------")
    for fam in families:
        if fam in shipped:
            where = "included in fonts/"
        elif fam in mac_only:
            where = "comes with macOS (see below)"
        elif fam.lower() in ON_WINDOWS:
            where = "comes with macOS and Windows"
        elif fam in kept_back:
            where = "NOT included (its license doesn't allow it): install it yourself"
        else:
            where = "not identified (see below)"
        w(f"  {fam}  ({chars_of(fam):,} characters): {where}")
        w(f"      styles: {styles_of(fam)}")
    w("")
    if mac_only:
        w("Mac fonts on other systems")
        w("  " + ", ".join(mac_only) + (" comes" if len(mac_only) == 1 else " come") + " with every Mac but can't")
        w("  be redistributed. On Windows, PowerPoint shows a similar font instead:")
        for fam in mac_only:
            w(f"    {fam} -> {WINDOWS_STANDIN.get(fam.lower(), 'a similar font')}")
        w("  Only small pieces of text use them here, so the difference is minor.")
        w("")
    if kept_back:
        w("Fonts you need to get yourself")
        w("  " + ", ".join(sorted(kept_back)) + ": the license doesn't allow passing the")
        w("  font file on. Get it from its maker (or your organisation) and install it as above.")
        w("")
    if missing:
        w("About \"not identified\" fonts")
        w("  The original didn't name these fonts and no installed font matched them, so")
        w("  the converter used a close substitute (Helvetica Neue or Menlo). They")
        w("  appear as those fonts in the presentation; nothing to install.")
        w("")

    w("Opening it in other apps")
    w("------------------------")
    w("  Keynote on iPad/iPhone   Fonts must be installed on the device too: AirDrop the")
    w("                           font files to it, or use a font-installer app.")
    w("  Keynote for iCloud       Uses only its own font list; others are replaced.")
    w("  PowerPoint for the web   Can't use fonts you install; it substitutes them.")
    w("                           Use the desktop app for the exact look.")
    w("  Google Slides            Upload the .pptx to Google Drive and open it with Google")
    w("                           Slides. Slides can't use installed fonts, but most open")
    w("                           fonts are on Google Fonts: in the font menu choose More")
    w("                           fonts, then search for each font listed above.")
    w("")

    w("Editing")
    w("-------")
    w("  - All text is real, editable text with the original fonts, sizes, spacing and colours.")
    if report.get("input") == "HTML deck":
        w("  - Photos are separate pictures made from the original image files: move,")
        w("    resize, re-crop or replace them. A photo with an overlay on it (a gradient,")
        w("    a caption plate) is grouped with it; double-click the group to reach the photo.")
        w("  - Cards, arrows, badges and backgrounds are separate pictures too.")
        w("  - Every object is named (\"Photo · ...\", \"Overlay\", \"Shape\", \"Background\").")
        w("    In Keynote, View > Show Object List lets you select anything, even objects")
        w("    underneath others.")
    else:
        w("  - Everything except the text (photos, shapes, backgrounds) is one picture per")
        w("    slide behind the text, exactly as in the original PDF.")
    w("")

    w("How faithful is it?")
    w("-------------------")
    w(f"  Converted from: {report.get('input', '')}" + (f" ({report['deck']})" if report.get("deck") else ""))
    w(f"  Slides: {report.get('slides', 0)}")
    if acc:
        worst = ", ".join(f"slide {p} ({m:.1f}%)" for p, m in acc["worst"])
        w(f"  Pixel match with the original: {acc['average']:.1f}% on average; lowest: {worst}.")
        w("  (Measured with all the fonts installed. Remaining differences are mostly")
        w("  anti-aliasing and image compression.)")
    else:
        w("  The slide-by-slide comparison needs Keynote and wasn't run for this export.")
    w("  Details: results/report.txt, and side-by-side images in results/slides/")
    w("  (original, converted, and a heatmap of the differences).")
    w("")

    w("If something looks off")
    w("----------------------")
    w("  Headings in the wrong font or weight")
    w("      A font isn't installed, or the app was open while you installed it.")
    w("      Install the fonts/ folder, quit the app completely, reopen.")
    w("  Keynote says \"fonts are missing\"")
    w("      Same cause and fix. Click Show to see which fonts it means.")
    w("  Text wraps onto an extra line in PowerPoint")
    w("      PowerPoint spaces lines slightly differently from Keynote. Widen the text")
    w("      box a little, or open the .key in Keynote for the exact layout.")
    w("  PowerPoint on Windows shows Arial or Consolas in places")
    w("      Those bits use Mac fonts (see above); this is expected.")
    w("")
    return "\n".join(L)


def write(folder, report, outputs, title, preview_dir=None):
    """Add the README, fonts and results next to the converted files in `folder`.
    outputs: {"key": path, "pptx": path} for the files already written there."""
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    usage = report.get("font_usage") or []
    shipped, kept_back = _ship_fonts(usage, folder / "fonts")
    if shipped:
        _license_file(shipped, folder / "fonts" / "FONTS-LICENSE.txt")

    results = folder / "results"
    results.mkdir(exist_ok=True)
    (results / "report.txt").write_text(reporting.as_text(report).lstrip() + "\n", encoding="utf-8")
    (results / "report.json").write_text(json.dumps(reporting.plain(report), indent=1), encoding="utf-8")
    if preview_dir and Path(preview_dir).exists() and Path(preview_dir).resolve() != (results / "slides").resolve():
        shutil.copytree(preview_dir, results / "slides", dirs_exist_ok=True)

    (folder / "README.txt").write_text(_readme(title, report, outputs, usage, shipped, kept_back), encoding="utf-8")
    return folder


def zip_folder(folder, zip_path):
    """The export folder as one .zip (for the web download)."""
    folder, zip_path = Path(folder), Path(zip_path)
    base = shutil.make_archive(str(zip_path.with_suffix("")), "zip", folder.parent, folder.name)
    return Path(base)
