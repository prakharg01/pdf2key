"""Command line: pdf2key input [output] [--format key|pptx|both]"""
import argparse
import sys
from pathlib import Path

from . import bundle, report
from .converter import convert, missing_fonts


def main(argv=None):
    ap = argparse.ArgumentParser(prog="pdf2key", description="Convert a PDF into an editable Keynote and/or "
                                 "PowerPoint presentation. Experimental: give an HTML deck (or a .zip of one) "
                                 "to also get movable pictures.")
    ap.add_argument("pdf", help="input PDF, or an HTML deck export: the .zip, the unzipped folder, "
                                "or the .html inside it (experimental)")
    ap.add_argument("output", nargs="?", help="export folder (default: '<name> - pdf2key' next to the input); "
                                              "with --no-bundle, the output file")
    ap.add_argument("--format", choices=("key", "pptx", "both"), default=None,
                    help="key, pptx (PowerPoint; skips Keynote, so it's fast and needs no Keynote), or both "
                         "(default: both, or key with --no-bundle)")
    ap.add_argument("--no-bundle", action="store_true",
                    help="write just the presentation file(s), without the export folder's README, fonts "
                         "and results")
    ap.add_argument("--pptx-only", action="store_true", help=argparse.SUPPRESS)  # old spelling of --format pptx
    ap.add_argument("--separate-images", action="store_true",
                    help="PDFs: make pictures individually movable (less faithful for cropped/masked images)")
    ap.add_argument("--no-verify", action="store_true", help="skip the comparison with the original")
    ap.add_argument("--deck", help="HTML exports holding several decks: which one (part of its file name)")
    ap.add_argument("--list-decks", action="store_true", help="list the decks in an HTML export and stop")
    args = ap.parse_args(argv)
    bundled = not (args.no_bundle or args.pptx_only)
    fmt = "pptx" if args.pptx_only else args.format or ("both" if bundled else "key")

    src = Path(args.pdf).resolve()
    if not src.exists():
        ap.error(f"{src} not found")
    stem = src.name if src.is_dir() else src.stem
    if bundled:
        folder = Path(args.output).resolve() if args.output else bundle.folder_for(src)
        if folder.suffix.lower() in (".key", ".pptx"):
            folder = bundle.folder_for(folder)
        base = folder / f"{stem}.key"
        preview_dir = folder / "results" / "slides"
    else:
        base = Path(args.output).resolve() if args.output else (
            src.parent / f"{src.name}.key" if src.is_dir() else src.with_suffix(".key"))
        preview_dir = None
    key_out = base.with_suffix(".key") if fmt in ("key", "both") else None
    pptx_out = base.with_suffix(".pptx") if fmt in ("pptx", "both") else None
    keynote = key_out is not None
    log = lambda msg, frac: print(f"  {msg}", file=sys.stderr)

    is_html = src.is_dir() or src.suffix.lower() in (".html", ".htm", ".zip")
    rep = report.new("HTML deck" if is_html else "PDF")
    if is_html:
        from .html import convert_html, decks_in
        decks = decks_in(src)
        if args.list_decks:
            for i, (name, n) in enumerate(decks):
                print(f"{'*' if i == 0 else ' '} {name}  ({n} slides)")
            return
        if len(decks) > 1 and not args.deck:
            print(f"  {len(decks)} decks in this export; converting {decks[0][0]!r} "
                  f"(choose another with --deck, see --list-decks)", file=sys.stderr)
        reports = convert_html(src, key_out, not args.no_verify, log, preview_dir, report=rep, deck=args.deck,
                               pptx_out=pptx_out, keynote=keynote)
    else:
        fonts = missing_fonts(src)
        if fonts:
            print(f"  fonts not installed, substituting: {', '.join(fonts)}", file=sys.stderr)
        reports = convert(src, key_out, args.separate_images, not args.no_verify, log, preview_dir, report=rep,
                          pptx_out=pptx_out, keynote=keynote)
    for r in reports:
        print(f"  slide {r.page}: {r.match:.2f}% match, {r.off_pixels:.2f}% of pixels differ")
    print(report.as_text(rep))
    if bundled:
        outputs = {k: v for k, v in (("key", key_out), ("pptx", pptx_out)) if v}
        bundle.write(folder, rep, outputs, Path(rep.get("deck") or stem).stem, preview_dir)
        print(f"Wrote {folder}/")
        for f in sorted(folder.iterdir()):
            print(f"  {f.name}{'/' if f.is_dir() else ''}")
        return
    for f in (key_out, pptx_out):
        if f:
            print(f"Wrote {f}")


if __name__ == "__main__":
    main()
