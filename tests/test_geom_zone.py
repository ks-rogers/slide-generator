"""非テキストプリミティブ（connector / picture / chart）の検査回帰テスト。

これらは描画文字を持たないため _FIT_REGISTRY ではなく _GEOM_REGISTRY に
bbox を記録し、validate_fit が本文コンテンツゾーン（x 1.07–24.33 /
y 3.0–12.8）からの逸脱を ERROR(ZONE) として検知する。生API（add_connector /
add_picture / add_chart を直接呼ぶ）と違い、これらのヘルパー経由なら検査ゲートに
乗ることを保証する。
"""
import pytest

import slides
from slides import (
    connector, picture, bar_chart, line_chart, pie_chart,
    validate_fit, load_template, L_BODY, SECONDARY, PRIMARY,
)


def _zone_kinds(prs):
    """検知結果のうち ZONE の (kind) 集合を返す。"""
    return {v.kind for v in validate_fit(prs) if v.code == "ZONE"}


@pytest.fixture
def sample_png(tmp_path):
    """テスト用の 200x100 PNG（2:1）を作って返す。"""
    from PIL import Image
    p = tmp_path / "sample.png"
    Image.new("RGB", (200, 100), (31, 78, 121)).save(p)
    return p


# ---------------------------------------------------------------- connector
def test_connector_in_zone_no_violation(prs, body_slide):
    connector(body_slide, 2.0, 4.0, 10.0, 4.0, color=SECONDARY, end_arrow=True)
    assert "connector" not in _zone_kinds(prs), [str(v) for v in validate_fit(prs)]


def test_connector_out_of_zone_is_zone_error(prs, body_slide):
    # 右端 30cm > X_MAX(24.33)
    connector(body_slide, 20.0, 4.0, 30.0, 4.0)
    assert "connector" in _zone_kinds(prs)


def test_connector_registers_bbox_as_minmax(prs, body_slide):
    """始点/終点の前後関係に依らず min/abs で矩形化される。"""
    connector(body_slide, 10.0, 9.0, 4.0, 4.0)  # 右上→左下
    rec = slides._GEOM_REGISTRY[-1]
    assert rec["kind"] == "connector"
    assert rec["x"] == pytest.approx(4.0)
    assert rec["y"] == pytest.approx(4.0)
    assert rec["w"] == pytest.approx(6.0)
    assert rec["h"] == pytest.approx(5.0)


# ---------------------------------------------------------------- picture
def test_picture_in_zone_no_violation(prs, body_slide, sample_png):
    picture(body_slide, sample_png, 12.0, 4.0, w=4.0)
    assert "picture" not in _zone_kinds(prs)


def test_picture_keeps_aspect_ratio(prs, body_slide, sample_png):
    """w のみ指定 → アスペクト比（2:1）を保って h が自動算出される。"""
    pic = picture(body_slide, sample_png, 12.0, 4.0, w=4.0)
    # 200x100 → 4.0cm 幅なら高さ 2.0cm
    assert slides.Emu(pic.width).cm == pytest.approx(4.0, abs=0.02)
    assert slides.Emu(pic.height).cm == pytest.approx(2.0, abs=0.02)


def test_picture_out_of_zone_is_zone_error(prs, body_slide, sample_png):
    # 下端 10.0 + 6.0 = 16cm > Y_MAX(12.8)
    picture(body_slide, sample_png, 2.0, 10.0, w=6.0, h=6.0)
    assert "picture" in _zone_kinds(prs)


def test_picture_missing_file_raises(body_slide, tmp_path):
    with pytest.raises(FileNotFoundError):
        picture(body_slide, tmp_path / "does_not_exist.png", 2.0, 4.0, w=3.0)


# ---------------------------------------------------------------- charts
def test_bar_chart_in_zone_no_violation(prs, body_slide):
    bar_chart(body_slide, 2.0, 5.0, 9.0, 6.0,
              categories=["A", "B", "C"],
              series=[("2024", [10, 20, 30]), ("2025", [15, 25, 35])],
              colors=[PRIMARY, SECONDARY], number_format="#,##0")
    assert "chart" not in _zone_kinds(prs)


def test_chart_out_of_zone_is_zone_error(prs, body_slide):
    # y=1.0 は見出し領域（Y_MIN=3.0 未満）
    bar_chart(body_slide, 2.0, 1.0, 9.0, 5.0,
              categories=["A"], series=[("s", [1])])
    assert "chart" in _zone_kinds(prs)


def test_line_chart_builds_and_registers(prs, body_slide):
    line_chart(body_slide, 2.0, 4.0, 10.0, 6.0,
               categories=["1月", "2月", "3月"],
               series=[("売上", [100, 120, 140])])
    assert slides._GEOM_REGISTRY[-1]["kind"] == "chart"
    assert "chart" not in _zone_kinds(prs)


def test_pie_chart_builds_and_registers(prs, body_slide):
    pie_chart(body_slide, 13.0, 4.0, 9.0, 6.0,
              categories=["X", "Y", "Z"], values=[50, 30, 20],
              show_percentage=True)
    assert slides._GEOM_REGISTRY[-1]["kind"] == "chart"
    assert "chart" not in _zone_kinds(prs)


# ---------------------------------------------------------------- registry reset
def test_load_template_clears_geom_registry(prs, body_slide, sample_png):
    connector(body_slide, 2.0, 4.0, 8.0, 4.0)
    picture(body_slide, sample_png, 10.0, 4.0, w=3.0)
    assert len(slides._GEOM_REGISTRY) >= 2
    load_template()  # 次の生成ラン開始時にレジストリをリセット
    assert slides._GEOM_REGISTRY == []
