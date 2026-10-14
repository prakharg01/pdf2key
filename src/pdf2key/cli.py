"""Command line: pdf2key input.pdf [output.key]"""
import argparse
import sys
from pathlib import Path

from .converter import convert, missing_fonts, pdf_to_pptx


def main(argv=None):
    ap = argparse.ArgumentParser(prog="pdf2key", description="Convert a PDF into an editable Keynote presentation.")
    ap.add_argument("pdf", help="input PDF")
    ap.add_argument("output", nargs="?", help="output .key (default: next to the PDF)")
    ap.add_argument("--pptx-only", action="store_true", help="write a .pptx and skip Keynote")
    ap.add_argument("--separate-images", action="store_true",
                    help="make pictures individually movable (less faithful for cropped/masked images)")
    ap.add_argument("--no-verify", action="store_true", help="skip the PDF-vs-Keynote comparison")
    args = ap.parse_args(argv)

    pdf = Path(args.pdf).resolve()
    if not pdf.exists():
        ap.error(f"{pdf} not found")
    out = Path(args.output).resolve() if args.output else pdf.with_suffix(".key")
    log = lambda msg, frac: print(f"  {msg}", file=sys.stderr)

    fonts = missing_fonts(pdf)
    if fonts:
        print(f"  fonts not installed, substituting: {', '.join(fonts)}", file=sys.stderr)

    if args.pptx_only:
        pptx = out.with_suffix(".pptx")
        pdf_to_pptx(pdf, pptx, args.separate_images, log)
        print(f"Wrote {pptx}")
        return
    reports = convert(pdf, out, args.separate_images, not args.no_verify, log)
    for r in reports:
        print(f"  slide {r.page}: {r.match:.2f}% match, {r.off_pixels:.2f}% of pixels differ")
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
