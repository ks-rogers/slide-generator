"""ARROW_TOO_SHORT（矢印潰れ）検知の回帰テスト。

矢じり付きコネクタは「総長 − 矢じり長 = 軸の可視長」が MIN_ARROW_SHAFT_CM を
下回ると、矢じりだけが潰れて向きが読めない矢印になる。validate_fit はこれを
WARN(ARROW_TOO_SHORT) で報告する（コミット 6f64ff6 で追加。geom のゾーン検査
では拾えない盲点）。

閾値・係数は slides の定数（MIN_ARROW_SHAFT_CM / ARROW_HEAD_K / PT_PER_CM）を
参照し、テストにハードコードしない。矢印は軸長＝座標差で決まるためフォント幅に
依存せず、CI（Arial 無し）で決定的に通る。
"""
import pytest

import slides
from slides import (
    connector, validate_fit,
    ARROW_HEAD_K, MIN_ARROW_SHAFT_CM, PT_PER_CM,
)


def _head_cm(width_pt):
    """線幅(pt)に対する矢じり 1 個分の長さ（cm）。実装（validate_fit）と同じ式。"""
    return ARROW_HEAD_K * width_pt / PT_PER_CM


def _arrow_violations(prs):
    return [v for v in validate_fit(prs) if v.code == "ARROW_TOO_SHORT"]


# ---------------------------------------------------------------- 登録経路
def test_no_arrow_connector_not_registered(prs, body_slide):
    """矢印なしコネクタは _ARROW_REGISTRY に載らず ARROW_TOO_SHORT を出さない。"""
    # 軸長 1.0cm（矢印が付いていれば太線で潰れる長さ）だが矢印なしなので無関係。
    connector(body_slide, 2.0, 6.0, 3.0, 6.0,
              begin_arrow=False, end_arrow=False)
    assert slides._ARROW_REGISTRY == []
    assert _arrow_violations(prs) == []


def test_register_arrow_records_minmax_bbox_and_fields(prs, body_slide):
    """始点/終点の前後関係に依らず bbox は min/abs 化し、length/width/heads を記録。"""
    connector(body_slide, 10.0, 9.0, 4.0, 5.0,  # 右上→左下（dx=-6, dy=-4）
              width=2.0, begin_arrow=True, end_arrow=True)
    rec = slides._ARROW_REGISTRY[-1]
    assert rec["x"] == pytest.approx(4.0)
    assert rec["y"] == pytest.approx(5.0)
    assert rec["w"] == pytest.approx(6.0)
    assert rec["h"] == pytest.approx(4.0)
    assert rec["length"] == pytest.approx((6.0 ** 2 + 4.0 ** 2) ** 0.5)
    assert rec["width"] == pytest.approx(2.0)
    assert rec["heads"] == 2  # begin + end の両端


# ---------------------------------------------------------------- 検知/無検知
def test_long_arrow_no_violation(prs, body_slide):
    """軸が十分長い矢印（可視軸 ≫ 下限）は無検知。"""
    connector(body_slide, 2.0, 6.0, 10.0, 6.0, end_arrow=True)  # 軸長 8cm
    assert _arrow_violations(prs) == []


def test_short_shaft_arrow_is_warn(prs, body_slide):
    """太線×短総長で可視軸 < MIN_ARROW_SHAFT_CM → WARN(ARROW_TOO_SHORT)。"""
    width = 10.0
    # 総長を下限値そのものに取ると、矢じり 1 個分（head_cm ≫ 下限）を引いた可視軸は
    # 確実に負＝下限割れになる。閾値は定数から取りハードコードしない。
    length = MIN_ARROW_SHAFT_CM
    connector(body_slide, 2.0, 6.0, 2.0 + length, 6.0,
              width=width, end_arrow=True)
    vs = _arrow_violations(prs)
    assert len(vs) == 1
    assert vs[0].severity == "WARN"
    assert vs[0].kind == "connector"


def test_heads_count_affects_shaft(prs, body_slide):
    """両端矢印（heads=2）は片端（heads=1）より軸が短く判定される。

    MIN + head ≤ length < MIN + 2*head となる総長を定数から導くと、片端なら
    可視軸 ≥ 下限で無検知、両端なら矢じり 2 個分を引いて下限割れ＝検知になる。
    head_cm 計算に heads が効いていることの確認。
    """
    width = 10.0
    head = _head_cm(width)
    length = MIN_ARROW_SHAFT_CM + 1.5 * head  # 上記範囲の中点
    # 片端（end のみ）→ 可視軸 = length - head ≥ 下限 → 無検知
    connector(body_slide, 2.0, 6.0, 2.0 + length, 6.0,
              width=width, end_arrow=True)
    assert _arrow_violations(prs) == []
    # 両端 → 可視軸 = length - 2*head < 下限 → 検知（総長は同じ）
    connector(body_slide, 2.0, 8.0, 2.0 + length, 8.0,
              width=width, begin_arrow=True, end_arrow=True)
    vs = _arrow_violations(prs)
    assert len(vs) == 1
    assert vs[0].severity == "WARN"
