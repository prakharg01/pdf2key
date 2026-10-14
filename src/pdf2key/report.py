"""The status report shown after a conversion: results, what was built, and
every change made to the fonts and text along the way."""
from collections import Counter


def new(kind):
    return {"input": kind, "slides": 0, "seconds": 0.0, "accuracy": None, "objects": {},
            "fonts": {"recognised": [], "installed": [], "substituted": [], "estimated": []},
            "text": {"boxes": 0, "letter_spaced": 0, "justified_lines": 0, "rotated": 0, "scales": []},
            "warnings": [], "identified": []}


def text_stats(report):
    return report["text"] if report is not None else None


def record_fonts(report, pages_blocks, span_font, anonymous):
    """How each piece of text's font was resolved, weighted by characters."""
    if report is None:
        return
    recognised, installed, substituted, estimated = Counter(), Counter(), Counter(), Counter()
    for blocks in pages_blocks:
        for lines in blocks:
            for line in lines:
                for s in line["spans"]:
                    n = len(s["text"].strip())
                    if not n:
                        continue
                    family, bold, italic, _, exact = span_font(s)
                    style = family + (" Bold" if bold else "") + (" Italic" if italic else "")
                    if s.get("face"):
                        recognised[style] += n
                    elif anonymous(s["font"]):
                        estimated[style] += n
                    elif exact:
                        installed[style] += n
                    else:
                        original = s["font"].split("+")[-1]
                        substituted[(original, style)] += n
    f = report["fonts"]
    f["recognised"] = recognised.most_common()
    f["installed"] = installed.most_common()
    f["substituted"] = [(a, b, n) for (a, b), n in substituted.most_common()]
    f["estimated"] = estimated.most_common()
    report["identified"] = [name for name, _ in f["recognised"]]


def record_objects(report, prs):
    """Count what ended up on the slides (by kind)."""
    if report is None:
        return
    counts = Counter()

    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:  # group
                counts["photo groups"] += 1
                walk(sh.shapes)
            elif sh.has_text_frame and sh.text_frame.text.strip():
                counts["text boxes"] += 1
            elif sh.shape_type == 13:
                name = sh.name.split(" ·")[0]
                counts[{"Photo": "photos", "Shape": "shapes", "Overlay": "overlays",
                        "Background": "backgrounds"}.get(name, "artwork layers")] += 1
    for slide in prs.slides:
        walk(slide.shapes)
    report["objects"] = dict(counts)
    report["slides"] = len(prs.slides)


def record_accuracy(report, slide_reports, seconds):
    if report is None:
        return
    report["seconds"] = round(seconds, 1)
    if slide_reports:
        worst = sorted(slide_reports, key=lambda r: r.match)[:3]
        report["accuracy"] = {"average": round(sum(r.match for r in slide_reports) / len(slide_reports), 2),
                              "worst": [(r.page, r.match) for r in worst]}


def as_text(report):
    """Plain-text version for the command line."""
    out = ["", "Conversion report", "-----------------",
           f"Input: {report['input']}" + (f" ({report['deck']})" if report.get("deck") else "")
           + f", {report['slides']} slides, {report['seconds']:.0f} s"]
    acc = report["accuracy"]
    if acc:
        worst = ", ".join(f"slide {p} ({m:.1f}%)" for p, m in acc["worst"])
        out.append(f"Match with the original: {acc['average']:.1f}% average; lowest: {worst}")
    if report["objects"]:
        out.append("Objects: " + ", ".join(f"{n} {k}" for k, n in sorted(report["objects"].items())))
    f = report["fonts"]
    if f["recognised"]:
        out.append("Fonts recognised from the glyphs (the PDF had no names for them):")
        out += [f"  {name}  ({n} characters)" for name, n in f["recognised"]]
    if f["installed"]:
        out.append("Fonts used as named in the original:")
        out += [f"  {name}  ({n} characters)" for name, n in f["installed"]]
    if f["substituted"]:
        out.append("Fonts not installed, substituted:")
        out += [f"  {a} -> {b}  ({n} characters)" for a, b, n in f["substituted"]]
    if f["estimated"]:
        out.append("Unnamed fonts with no installed match, style estimated:")
        out += [f"  {name}  ({n} characters)" for name, n in f["estimated"]]
    t = report["text"]
    changes = []
    if t["scales"]:
        lo, hi = min(t["scales"]), max(t["scales"])
        changes.append(f"{len(t['scales'])} text boxes resized {(lo - 1) * 100:+.0f}% to {(hi - 1) * 100:+.0f}% "
                       f"so substituted fonts keep the original line widths")
    if t["letter_spaced"]:
        changes.append(f"{t['letter_spaced']} runs given the original letter-spacing")
    if t["justified_lines"]:
        changes.append(f"{t['justified_lines']} justified lines kept justified by widening their spaces")
    if t["rotated"]:
        changes.append(f"{t['rotated']} rotated text boxes")
    if changes:
        out.append(f"Text ({t['boxes']} boxes):")
        out += [f"  {c}" for c in changes]
    for w in report["warnings"]:
        out.append(f"Warning: {w}")
    return "\n".join(out)
