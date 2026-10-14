"""Tests for the PDF -> .pptx stage (no Keynote needed)."""
import pymupdf
import pytest
from pptx import Presentation

from pdf2key import pdf_to_pptx
from pdf2key.converter import _font_info


@pytest.fixture
def sample_pdf(tmp_path):
    doc = pymupdf.open()
    page = doc.new_page(width=800, height=450)
    page.insert_text((50, 80), "Title", fontname="tiro", fontsize=40)
    page.insert_text((50, 140), "Line one", fontsize=18, color=(1, 0, 0))
    page.insert_text((50, 162), "Line two", fontsize=18, color=(1, 0, 0))
    page.insert_text((400, 300), "Rotated", fontsize=20, rotate=90)
    page.draw_line((50, 200), (700, 200), color=(0, 0, 1), width=2)
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 40, 40), False)
    pix.set_rect(pix.irect, (0, 200, 0))
    page.insert_image(pymupdf.Rect(600, 300, 700, 400), pixmap=pix)
    portrait = doc.new_page(width=450, height=800)
    portrait.insert_text((50, 100), "Portrait", fontsize=30)
    path = tmp_path / "sample.pdf"
    doc.save(path)
    return path


def _texts(slide):
    return [s.text_frame.text for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]


def test_builds_one_slide_per_page(sample_pdf, tmp_path):
    out = tmp_path / "out.pptx"
    pdf_to_pptx(sample_pdf, out)
    prs = Presentation(out)
    assert len(prs.slides) == 2
    assert prs.slide_width == 800 * 12700 and prs.slide_height == 450 * 12700


def test_text_is_editable_and_grouped(sample_pdf, tmp_path):
    out = tmp_path / "out.pptx"
    pdf_to_pptx(sample_pdf, out)
    texts = _texts(Presentation(out).slides[0])
    assert "Title" in texts
    assert "Line one\nLine two" in texts  # consecutive lines stay one text box
    assert "Rotated" in texts


def test_rotated_text_keeps_rotation(sample_pdf, tmp_path):
    out = tmp_path / "out.pptx"
    pdf_to_pptx(sample_pdf, out)
    box = next(s for s in Presentation(out).slides[0].shapes if s.has_text_frame and s.text_frame.text == "Rotated")
    assert round(box.rotation) % 90 == 0 and round(box.rotation) != 0


def test_images_and_line_art_are_pictures(sample_pdf, tmp_path):
    out = tmp_path / "out.pptx"
    pdf_to_pptx(sample_pdf, out)
    pictures = [s for s in Presentation(out).slides[0].shapes if s.shape_type == 13]
    assert len(pictures) == 2  # the embedded image + the line-art overlay

    pdf_to_pptx(sample_pdf, out, background=False)
    pictures = [s for s in Presentation(out).slides[0].shapes if s.shape_type == 13]
    assert len(pictures) == 1


def test_smaller_page_is_fit_inside_slide(sample_pdf, tmp_path):
    out = tmp_path / "out.pptx"
    pdf_to_pptx(sample_pdf, out)
    prs = Presentation(out)
    for shape in prs.slides[1].shapes:
        assert 0 <= shape.left and shape.left + shape.width <= prs.slide_width + 12700 * 8
        assert 0 <= shape.top and shape.top + shape.height <= prs.slide_height


@pytest.mark.parametrize("name, family_kind", [
    ("ABCDEF+SomeUnknownMono-Bold", "Menlo"),
    ("XYZ+UnknownSerifPro-Regular", "Times New Roman"),
    ("Q+TotallyUnknownSans", "Helvetica Neue"),
])
def test_font_fallbacks(name, family_kind):
    family, *_ , exact = _font_info(name, 0)
    assert not exact and family == family_kind


def test_bold_italic_detection():
    _, bold, italic, *_ = _font_info("ABC+Unknown-BoldItalic", 0)
    assert bold and italic
