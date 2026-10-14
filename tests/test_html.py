"""Tests for the experimental HTML deck input."""
import pytest
from PIL import Image
from pptx import Presentation

from pdf2key.html import _displayed_rect

EMU_PER_PX = 12700 * 0.75


def test_object_fit_cover_crops_the_overflow():
    # 400x200 image in a 100x100 box, cover, centred: drawn 200x100, 50px off each side
    item = {"kind": "image", "natural": [400, 200], "fit": "cover", "pos": "50% 50%",
            "content": {"x": 0, "y": 0, "w": 100, "h": 100}}
    assert _displayed_rect(item) == pytest.approx((-50, 0, 200, 100))


def test_background_size_contain_and_position():
    item = {"kind": "bgimage", "natural": [200, 100], "size": "contain", "pos": "0% 100%",
            "rect": {"x": 10, "y": 10, "w": 100, "h": 100}}
    assert _displayed_rect(item) == pytest.approx((10, 60, 100, 50))


@pytest.fixture
def chrome():
    pw = pytest.importorskip("playwright.sync_api")
    try:
        with pw.sync_playwright() as p:
            p.chromium.launch(channel="chrome").close()
    except Exception:
        pytest.skip("Google Chrome not available")


def test_html_deck_gets_movable_cropped_pictures_under_overlays(chrome, tmp_path):
    from pdf2key.html import html_to_pptx
    Image.new("RGB", (800, 400), (200, 40, 40)).save(tmp_path / "photo.png")
    (tmp_path / "deck.html").write_text("""<!doctype html><html><body>
      <section style="position:relative;width:1920px;height:1080px;background:#111">
        <img src="photo.png" style="position:absolute;left:100px;top:100px;width:400px;height:400px;object-fit:cover">
        <div style="position:absolute;left:100px;top:100px;width:400px;height:400px;
                    background:linear-gradient(transparent,#000)"></div>
        <h1 style="position:absolute;left:600px;top:100px;color:#fff;font:64px Helvetica">Hello deck</h1>
      </section>
      <section style="position:relative;width:1920px;height:1080px;background:#fff"><p>Second</p></section>
    </body></html>""")
    out = tmp_path / "deck.pptx"
    html_to_pptx(tmp_path / "deck.html", out, tmp_path / "deck.pdf")
    slides = Presentation(out).slides
    assert len(slides) == 2
    shapes = list(slides[0].shapes)
    pics = [s for s in shapes if s.shape_type == 13]
    photo = next(p for p in pics if p.crop_left > 0)          # the <img>, cropped by object-fit
    assert photo.crop_left == pytest.approx(0.25, abs=0.01) and photo.crop_right == pytest.approx(0.25, abs=0.01)
    assert photo.left == pytest.approx(100 * EMU_PER_PX, rel=0.01)
    overlay = shapes.index(photo) + 1                          # the gradient is a layer above it
    assert shapes[overlay].shape_type == 13
    assert any(s.has_text_frame and "Hello deck" in s.text_frame.text for s in shapes)
