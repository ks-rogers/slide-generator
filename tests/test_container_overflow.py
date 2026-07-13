"""CONTAINER_OVERFLOW（コンテナはみ出し）検知の回帰テスト。

register_container で登録した背景枠（カード/パネル）に対し、validate_fit は
「中心が枠内・枠より小さい子要素」を集め、子の bbox が枠を越えていれば
WARN(CONTAINER_OVERFLOW) を報告する（コミット 6f64ff6 で追加。geom のページ
ゾーン検査では拾えない「枠内のはみ出し」を補完する）。

スラック閾値は slides.CONTAINER_SLACK_CM を参照しハードコードしない。子要素の
判定は座標（bbox）で決まるためフォント幅に依存せず、CI（Arial 無し）で決定的。
テキスト溢れ等の他検知が混ざらないよう、検証は code=="CONTAINER_OVERFLOW" で絞る。
"""
import pytest

import slides
from slides import (
    textbox, picture, connector, card, validate_fit,
    register_container, L_BODY, CONTAINER_SLACK_CM,
)

# 標準コンテナ枠（本文コンテンツゾーン内）。cx1=12.0 / cy1=10.0 / 面積=60。
CX, CY, CW, CH = 2.0, 4.0, 10.0, 6.0
CX1, CY1 = CX + CW, CY + CH


@pytest.fixture
def sample_png(tmp_path):
    """テスト用の 200x100 PNG（2:1）を作って返す。"""
    from PIL import Image
    p = tmp_path / "sample.png"
    Image.new("RGB", (200, 100), (31, 78, 121)).save(p)
    return p


def _container_violations(prs):
    return [v for v in validate_fit(prs) if v.code == "CONTAINER_OVERFLOW"]


# ---------------------------------------------------------------- 無検知（健全）
def test_child_fully_inside_no_violation(prs, body_slide):
    """枠内に完全に収まる子要素は無検知。"""
    register_container(body_slide, CX, CY, CW, CH)
    textbox(body_slide, 3.0, 5.0, 4.0, 1.0, "枠内")
    assert _container_violations(prs) == []


# ---------------------------------------------------------------- 検知（4辺）
def test_child_overflows_right_and_bottom(prs, body_slide):
    """枠の右/下に飛び出す子要素 → WARN(CONTAINER_OVERFLOW)、detail に正しい辺。"""
    register_container(body_slide, CX, CY, CW, CH)
    # 右はみ出し: x+w=13 > cx1=12（中心 x=11.5 は枠内）
    textbox(body_slide, 10.0, 5.0, 3.0, 1.0, "右へ")
    # 下はみ出し: y+h=11 > cy1=10（中心 y=10.0 は枠内）
    textbox(body_slide, 3.0, 9.0, 3.0, 2.0, "下へ")
    vs = _container_violations(prs)
    assert len(vs) == 2
    assert all(v.severity == "WARN" for v in vs)
    details = " ".join(v.detail for v in vs)
    assert "右" in details and "下" in details


def test_child_overflows_left_and_top(prs, body_slide):
    """枠の左/上に飛び出す子要素 → 検知。"""
    register_container(body_slide, CX, CY, CW, CH)
    # 左はみ出し: ex=1.0 < cx0=2.0（中心 x=2.25 は枠内）
    textbox(body_slide, 1.0, 5.0, 2.5, 1.0, "左へ")
    # 上はみ出し: ey=3.0 < cy0=4.0（中心 y=4.0 は枠内）
    textbox(body_slide, 4.0, 3.0, 2.0, 2.0, "上へ")
    vs = _container_violations(prs)
    assert len(vs) == 2
    details = " ".join(v.detail for v in vs)
    assert "左" in details and "上" in details


# ---------------------------------------------------------------- 子の判定条件
def test_center_outside_frame_not_a_child(prs, body_slide):
    """中心が枠外の要素は子とみなさず無検知（cx0<=ecx<=cx1 条件）。"""
    register_container(body_slide, CX, CY, CW, CH)
    # 中心 x=13.0 > cx1=12 → 枠の子ではない（右に飛び出すが対象外）
    textbox(body_slide, 11.5, 5.0, 3.0, 1.0, "枠外中心")
    assert _container_violations(prs) == []


def test_element_as_large_as_frame_not_a_child(prs, body_slide):
    """枠と同等以上の大きさ（ew*eh >= c_area）の要素は子とみなさない。"""
    register_container(body_slide, CX, CY, CW, CH)
    # 面積 10.9*6.5=70.85 >= 60。中心(6.55,6.75)は枠内だが背景枠級なので除外。
    textbox(body_slide, 1.1, 3.5, 10.9, 6.5, "大枠")
    assert _container_violations(prs) == []


def test_overflow_within_slack_no_violation(prs, body_slide):
    """CONTAINER_SLACK_CM 内のわずかな超過は無検知（境界値）。"""
    register_container(body_slide, CX, CY, CW, CH)
    # 右に CONTAINER_SLACK_CM の半分だけ超過 → スラック内なので無検知。
    w = (CX1 - 10.0) + CONTAINER_SLACK_CM / 2
    textbox(body_slide, 10.0, 5.0, w, 1.0, "微超過")
    assert _container_violations(prs) == []


# ---------------------------------------------------------------- スライド境界
def test_other_slide_element_not_counted(prs, body_slide):
    """別スライドの要素は対象外（slide_id 一致条件）。"""
    register_container(body_slide, CX, CY, CW, CH)
    other = prs.slides.add_slide(prs.slide_layouts[L_BODY])
    # 同座標なら同スライドでは右はみ出し検知されるが、別スライドなので対象外。
    textbox(other, 10.0, 5.0, 3.0, 1.0, "別スライド")
    assert _container_violations(prs) == []


# ---------------------------------------------------------------- geom 子要素
def test_geom_children_are_collected(prs, body_slide, sample_png):
    """geom 要素（picture / connector）も子として拾われ、はみ出しを検知する。"""
    register_container(body_slide, CX, CY, CW, CH)
    # picture: x+w=13 > cx1=12（中心 x=11.5 枠内）→ 右はみ出し
    picture(body_slide, sample_png, 10.0, 5.0, w=3.0, h=1.0)
    # connector: 終点 x=13 > cx1=12（中心 x=11.5 枠内）→ 右はみ出し
    connector(body_slide, 10.0, 7.0, 13.0, 7.0)
    kinds = {v.kind for v in _container_violations(prs)}
    assert "picture" in kinds
    assert "connector" in kinds


# ---------------------------------------------------------------- card 統合
def test_card_registers_container(prs, body_slide):
    """card() が内部で register_container を呼ぶ（カードからはみ出すと検知）。"""
    card(body_slide, 2.0, 4.0, 8.0, 4.0, title="見出し", body="本文")  # cx1=10/cy1=8
    # カード右端(10.0)を越える子テキスト（中心 x=9.5 は枠内）
    textbox(body_slide, 8.0, 5.0, 3.0, 1.0, "カード外へ")
    vs = _container_violations(prs)
    assert len(vs) == 1
    assert vs[0].severity == "WARN"
