"""pdf2key: turn any PDF into an editable Keynote presentation."""
from .converter import SlideReport, convert, missing_fonts, pdf_to_pptx

__version__ = "0.1.0"
__all__ = ["SlideReport", "convert", "missing_fonts", "pdf_to_pptx", "__version__"]
