"""サンプル案件: サンプル商事株式会社向け AIチャットボット導入提案。

公開リポジトリ同梱のリファレンス実装。`content.txt`（架空の原稿）から
`/slides-gen` 相当の設計でスライドを組み立てる例を示す。

実行:
    source .venv/bin/activate
    python3 examples/サンプル商事_AIチャットボット/generate.py

末尾の validate() がレイアウト検査ゲートを通す。ERROR が出たら座標・サイズ・
テキスト量を調整して PASS させること（README / CLAUDE.md の検査ゲート節を参照）。
"""
import sys
from pathlib import Path

# examples/<案件名>/generate.py → リポジトリルートは parents[2]
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from slides import (
    PRIMARY, SECONDARY, SUCCESS, DANGER, TEXT, WHITE, T_BODY,
    L_BODY, L_CHAPTER,
    textbox, card, callout, connector, bar_chart,
    configure_body, configure_chapter, update_cover,
    load_template, reset_to_cover_only, add_back_cover,
    finalize_page_numbers, validate,
)

OUT = Path(__file__).parent / "output.pptx"

# ---- コンテンツゾーン基準（CLAUDE.md の雛形仕様より） ----
X0 = 1.07               # 左端
ZONE_W = 23.26          # 有効コンテンツ幅（右端 24.33）
LEAD_Y = 3.02           # タイトル下リード文の開始 y


def lead(s, text):
    """タイトル下のリード文（1行サマリ）。本文色 TEXT で描く（design-guide §6）。"""
    textbox(s, X0, LEAD_Y, ZONE_W, 0.62, text, size=11.5, color=TEXT)


def body(prs, *, section_label, title):
    """本文スライドを1枚追加して返す。"""
    s = prs.slides.add_slide(prs.slide_layouts[L_BODY])
    configure_body(s, section_label=section_label, title=title)
    return s


def chapter(prs, *, num, title, subtitle):
    """セクション扉スライドを1枚追加して返す。"""
    s = prs.slides.add_slide(prs.slide_layouts[L_CHAPTER])
    configure_chapter(s, chapter_num=num, title=title, subtitle=subtitle)
    return s


# ============================================================ 各ページ
def page_overview(prs):
    s = body(prs, section_label="OVERVIEW / 全体コンセプト",
             title="社内問い合わせを AI が一次対応する")
    lead(s, "「探す・聞く・待つ」をなくし、社員が本来業務に集中できる状態をつくる。")
    callout(s, X0, 3.85, ZONE_W, 1.1,
            "散在する社内ナレッジを AI に集約し、問い合わせの一次対応を自動化する。")

    cw = (ZONE_W - 2 * 0.6) / 3        # 3カラム等幅＋ガター0.6
    cy, ch = 5.35, 5.8
    cards = [
        ("1", "対応の属人化",
         "規程や手続きの問い合わせが特定担当に集中し、回答待ちが常態化している。", PRIMARY),
        ("2", "AIで一次対応",
         "既存のマニュアル・FAQ・規程を AI が参照し、24時間その場で一次回答する。", SECONDARY),
        ("3", "工数と時間を削減",
         "一次対応の自動化で、回答待ち時間と担当者の対応件数を同時に圧縮する。", SUCCESS),
    ]
    for i, (chip, title, txt, accent) in enumerate(cards):
        x = X0 + i * (cw + 0.6)
        card(s, x, cy, cw, ch, chip=chip, title=title, body=txt, accent=accent)


def page_background(prs):
    s = body(prs, section_label="BACKGROUND / 現状の課題",
             title="問い合わせ対応が人に依存したまま放置されている")
    lead(s, "いまの運用は「人に依存した一次対応」が前提で、滞りとばらつきを生んでいる。")

    cw = (ZONE_W - 0.7) / 2
    cy, ch = 3.95, 7.0
    card(s, X0, cy, cw, ch, title="AS-IS：いまの状態", kicker="人に依存した運用",
         accent=DANGER, body=[
             "・問い合わせ窓口が特定担当に集中し、不在時は回答が止まる",
             "・同じ質問が繰り返し寄せられ、都度ゼロから回答している",
             "・回答の根拠が個人の記憶に依存し、内容にばらつきがある",
         ])
    card(s, X0 + cw + 0.7, cy, cw, ch, title="TO-BE：あるべき姿",
         kicker="AIが一次対応", accent=SUCCESS, body=[
             "・一次対応を AI が担い、担当者は例外対応に専念できる",
             "・よくある質問はその場で即時回答され、待ち時間が消える",
             "・回答は常に最新のナレッジを根拠とし、品質が安定する",
         ])


def page_solution(prs):
    s = body(prs, section_label="SOLUTION / 提案の仕組み",
             title="「集める → 学習する → 答える」で自動化する")
    lead(s, "散在するナレッジを集約し、AI が根拠付きで一次回答する3ステップ。")

    cw, gap = 6.95, 1.2
    cy, ch = 4.2, 5.0
    steps = [
        ("1", "集める", "マニュアル・FAQ・規程など散在するナレッジを一箇所に集約する。"),
        ("2", "学習する", "集約したナレッジを AI が参照できる形に整え、引用可能にする。"),
        ("3", "答える", "社員はチャットで質問し、AI が根拠付きで一次回答する。"),
    ]
    xs = [X0 + i * (cw + gap) for i in range(3)]
    for x, (chip, title, txt) in zip(xs, steps):
        card(s, x, cy, cw, ch, chip=chip, title=title, body=txt)

    # ステップ間を矢印でつなぐ（カードの縦中央で接続）
    ymid = cy + ch / 2
    for i in range(2):
        x1 = xs[i] + cw
        x2 = xs[i + 1]
        connector(s, x1 + 0.15, ymid, x2 - 0.15, ymid,
                  color=PRIMARY, width=1.5, end_arrow=True)

    callout(s, X0, 9.9, ZONE_W, 1.3,
            "社員はチャットで質問するだけ。AI が最新ナレッジを根拠に一次回答する。")


def page_outcome(prs):
    s = body(prs, section_label="OUTCOME / 期待効果",
             title="待ち時間と対応工数を同時に削減できる")
    lead(s, "一次対応の自動化で、回答待ち時間と担当者の対応工数を大きく削減できる。")

    # 左: 棒グラフ（導入前後の対応件数）。グラフ自動タイトルは切り、左に小見出しを置く。
    textbox(s, X0, 3.95, 13.0, 0.55, "担当者の月間対応件数（件 ／ 試算）",
            size=T_BODY, bold=True, color=TEXT)
    chart = bar_chart(s, X0, 4.6, 13.0, 6.0,
                      categories=["導入前", "導入後"],
                      series=[("対応件数", [320, 90])],
                      colors=[PRIMARY], number_format="#,##0")
    chart.has_title = False

    # 右: KPI カード2枚（見出し＋「変化（前→後）」＋補足）
    rx, rw = 14.5, 9.83
    card(s, rx, 3.95, rw, 3.2, title="回答待ち時間", accent=SECONDARY,
         body=["平均 4.0時間 → 0.5時間", "一次回答が即時化し待ち時間が消える"])
    card(s, rx, 7.4, rw, 3.2, title="月間対応件数", accent=SUCCESS,
         body=["320件 → 90件", "担当者の負荷を約7割削減できる"])

    callout(s, X0, 11.0, ZONE_W, 1.25,
            "AI が一次対応を担うことで、担当者は判断が必要な例外対応に集中できる。")


# ============================================================ 組み立て
def main():
    prs = load_template()

    # 表紙（テンプレ既存スライドを書き換え。add_slide では作らない）
    update_cover(prs.slides[0], lines=[
        {"text": "サンプル商事株式会社 御中", "size": 12, "color": WHITE},
        {"text": "社内問い合わせ対応 AIチャットボット導入のご提案",
         "size": 17, "bold": True, "color": WHITE},
        {"text": "回答待ち時間と担当者の工数を同時に削減します。",
         "size": 10, "italic": True, "color": WHITE},
        {"text": "株式会社サンプルソリューションズ　／　2026年6月　／　Ver. 1.0",
         "size": 10, "color": WHITE},
    ])
    reset_to_cover_only(prs)

    page_overview(prs)
    chapter(prs, num="PART  01", title="課題の整理",
            subtitle="いまの問い合わせ対応の何がボトルネックか")
    page_background(prs)
    chapter(prs, num="PART  02", title="ご提案内容",
            subtitle="社内ナレッジを AI に集約し対応を自動化する")
    page_solution(prs)
    page_outcome(prs)

    add_back_cover(prs)
    finalize_page_numbers(prs, skip_first=True)
    prs.save(OUT)
    print(f"saved: {OUT}  (slides: {len(prs.slides)})")
    validate(prs, OUT, render=True)


if __name__ == "__main__":
    main()
