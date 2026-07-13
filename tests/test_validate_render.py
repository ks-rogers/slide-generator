"""validate_render（レンダ後突合）のユニットテスト。

実レンダラ（soffice / PyMuPDF）は CI に無いため、`fitz` と
`_soffice_to_pdf` をモックし「描画結果がこうだったら何を検知するか」を
決定的に検証する。従来この確定検知層にはテストが無く、
「一部見切れ（CLIP）は捕まえるのに完全消失は無言 PASS」という逆転や
レンダ PDF のページ欠落の黙殺を見逃していた（公開前レビュー指摘）。
"""
import sys
import types

import pytest
from pptx.enum.shapes import MSO_SHAPE

import slides
from slides import shape_box, textbox, PRIMARY, PT_PER_CM, WHITE

INTENDED = "検査対象の意図テキスト"


def _span(text, x0, y0, x1, y1):
    """cm 座標指定の span を fitz が返す pt 座標 dict に変換する。"""
    return {"text": text, "bbox": [v * PT_PER_CM for v in (x0, y0, x1, y1)]}


def _fake_fitz(pages):
    """ページごとの span リストをそのまま返す fitz モジュールのモック。"""
    class Page:
        def __init__(self, spans):
            self._spans = spans

        def get_text(self, _kind):
            if not self._spans:
                return {"blocks": []}
            return {"blocks": [{"lines": [{"spans": self._spans}]}]}

    class Doc:
        def __iter__(self):
            return iter([Page(sp) for sp in pages])

        def close(self):
            pass

    class Tools:
        def mupdf_display_errors(self, _flag):
            pass

    mod = types.ModuleType("fitz")
    mod.open = lambda _path: Doc()
    mod.TOOLS = Tools()
    return mod


@pytest.fixture
def run_render(monkeypatch, prs):
    """最終ページの偽描画結果を与えて validate_render を実行するヘルパー。

    drop_last_page=True でレンダ PDF の最終ページ欠落を再現する。
    """
    monkeypatch.setattr(slides, "_soffice_to_pdf", lambda p, **k: "dummy.pdf")

    def _run(last_page_spans, *, drop_last_page=False):
        pages = [[] for _ in range(len(prs.slides))]
        pages[-1] = last_page_spans
        if drop_last_page:
            pages = pages[:-1]
        monkeypatch.setitem(sys.modules, "fitz", _fake_fitz(pages))
        return slides.validate_render(prs, "dummy.pptx")

    return _run


def _codes(violations):
    return [v.code for v in violations]


# ----------------------------------------------------------------- 正常系
def test_exact_render_passes(run_render, body_slide):
    textbox(body_slide, 1.07, 4.0, 20.0, 0.6, INTENDED, size=11)
    assert run_render([_span(INTENDED, 1.1, 4.05, 8.0, 4.55)]) == []


# ----------------------------------------------------------- CLIP（一部見切れ）
def test_partial_render_is_clip_error(run_render, body_slide):
    textbox(body_slide, 1.07, 4.0, 20.0, 0.6, INTENDED, size=11)
    v = run_render([_span(INTENDED[:5], 1.1, 4.05, 4.0, 4.55)])
    assert _codes(v) == ["CLIP"]
    assert v[0].severity == "ERROR"


# ----------------------------------------------------------- LOST（完全消失）
def test_vanished_text_is_lost_error(run_render, body_slide):
    """1 文字も描画されなかった要素は無言 PASS ではなく ERROR(LOST)。"""
    textbox(body_slide, 1.07, 4.0, 20.0, 0.6, INTENDED, size=11)
    v = run_render([])
    assert _codes(v) == ["LOST"]
    assert v[0].severity == "ERROR"


# ------------------------------------------------- PAGE_MISSING（ページ欠落）
def test_missing_pdf_page_is_error(run_render, body_slide):
    """レンダ PDF がスライド数より少ない場合は黙殺せず ERROR にする。"""
    textbox(body_slide, 1.07, 4.0, 20.0, 0.6, INTENDED, size=11)
    v = run_render([], drop_last_page=True)
    assert _codes(v) == ["PAGE_MISSING"]
    assert v[0].severity == "ERROR"


# --------------------------------------------------------- ZONE（実測ゾーン逸脱）
def test_rendered_bbox_out_of_zone_is_zone_error(run_render, body_slide):
    # 枠は下端ぎりぎり、実描画 bbox が Y_MAX(12.8) を超過するケース
    textbox(body_slide, 1.07, 12.4, 20.0, 0.6, INTENDED, size=11)
    v = run_render([_span(INTENDED, 1.1, 12.9, 8.0, 13.4)])
    assert "ZONE" in _codes(v)


# ------------------------------------------- SHAPE_OVERFLOW（親矩形からの逸脱）
def test_rendered_bbox_outside_parent_shape_is_error(run_render, body_slide):
    shape_box(body_slide, MSO_SHAPE.RECTANGLE, 2.0, 4.0, 6.0, 1.0,
              text=INTENDED, fill=PRIMARY, color=WHITE, size=11)
    # 意図テキストは全文描画されたが、bbox が親矩形の左端からはみ出す
    v = run_render([_span(INTENDED, 1.5, 4.2, 7.0, 4.7)])
    assert "SHAPE_OVERFLOW" in _codes(v)
