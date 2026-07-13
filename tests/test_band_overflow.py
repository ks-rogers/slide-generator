"""BAND_OVERFLOW 検査の回帰テスト。

「色帯の下端をテキストが突き抜ける」崩れを ERROR に格上げする検査が、
- 低い色帯を突き抜けるケースを ERROR(BAND_OVERFLOW) にする
- 帯内に収まるケースは無検知
- 背の高い白カード内の保守的な縦伸長（P3 型の空振り）は ERROR にせず WARN のまま
を満たすことを確認する。
"""
from pptx.enum.shapes import MSO_SHAPE

from slides import shape_box, textbox, validate_fit, SECONDARY, WHITE, TEXT


def _codes(prs):
    return [(v.severity, v.code) for v in validate_fit(prs)]


def test_text_leaks_out_of_low_band_is_error(prs, body_slide):
    """1.0cm の色帯に size13 の見出しを置くと帯下端を突き抜ける → ERROR。"""
    y0 = 3.75
    shape_box(body_slide, MSO_SHAPE.RECTANGLE, 1.07, y0, 7.5, 1.0,
              fill=SECONDARY, line=None)
    # size13 は約 0.70cm 必要。枠 0.55 で帯(下端 y0+1.0)を突き抜ける配置。
    textbox(body_slide, 1.37, y0 + 0.48, 6.9, 0.55, "端末紛失・盗難",
            size=13, bold=True, color=WHITE)
    assert ("ERROR", "BAND_OVERFLOW") in _codes(prs), _codes(prs)


def test_text_fits_in_band_no_error(prs, body_slide):
    """帯高 1.5cm に余裕を持って収めれば BAND_OVERFLOW は出ない。"""
    y0 = 3.75
    shape_box(body_slide, MSO_SHAPE.RECTANGLE, 1.07, y0, 7.5, 1.5,
              fill=SECONDARY, line=None)
    textbox(body_slide, 1.37, y0 + 0.66, 6.9, 0.72, "端末紛失・盗難",
            size=13, bold=True, color=WHITE)
    assert not [c for c in _codes(prs) if c[1] == "BAND_OVERFLOW"], _codes(prs)


def test_tall_white_card_conservative_overflow_stays_warn(prs, body_slide):
    """背の高い白カード内で need_h が枠を超えるだけ（P3 型）は BAND_OVERFLOW に
    せず、従来どおり OVERFLOW_V の WARN に留める。

    行数×サイズだけで溢れる短行を並べ、折返し（フォント幅依存）に頼らず
    どの環境でも決定的に縦あふれさせる（CI に Arial が無くても結果不変）。
    """
    y0 = 3.75
    # 白カード（高さ 8.7cm > BAND_MAX_H）。帯ではない。
    shape_box(body_slide, MSO_SHAPE.RECTANGLE, 1.07, y0, 7.5, 8.7,
              fill=WHITE, line=None)
    # 6 短行（各行は折返さない）。size10.5×ls1.4 で need_h ≈ 4.1cm > 枠 2.5cm。
    # 下端がコンテンツゾーン内（≤12.8cm）に収まる位置に置き、ZONE 逸脱を混ぜない。
    body = [f"防御レイヤー {i}" for i in range(1, 7)]
    textbox(body_slide, 1.52, y0 + 2.25, 6.6, 2.5, body,
            size=10.5, color=TEXT, line_spacing=1.4)
    codes = _codes(prs)
    assert not [c for c in codes if c[1] == "BAND_OVERFLOW"], codes
    assert ("WARN", "OVERFLOW_V") in codes, codes


def test_text_below_band_not_contained(prs, body_slide):
    """色帯より下に置かれた textbox（帯に内包されない）は対象外。"""
    y0 = 3.75
    shape_box(body_slide, MSO_SHAPE.RECTANGLE, 1.07, y0, 7.5, 0.8,
              fill=SECONDARY, line=None)
    # 帯(下端 y0+0.8)よりさらに下の位置に、わざと溢れる小枠で配置。
    textbox(body_slide, 1.37, y0 + 3.95, 6.9, 0.4, "端末の鍵だけが守る対象",
            size=17, bold=True, color=SECONDARY)
    assert not [c for c in _codes(prs) if c[1] == "BAND_OVERFLOW"], _codes(prs)
