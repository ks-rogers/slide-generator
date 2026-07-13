"""brand.py — ブランド設定の単一ソース（ここを編集すればリブランドできる）。

色・フォント・既定雛形ファイル名をこの1ファイルに集約している。`slides.py` が
これを読み込んで再エクスポートするため、generate.py やスキルは
`from slides import PRIMARY, ...` で参照する。

定数名は「色相」ではなく「役割」で名付けてある（`PRIMARY` `SECONDARY`
`SUCCESS` `DANGER` …）。リブランドは **値（RGBColor）だけ** を差し替え、名前は
据え置く — こうすれば `PRIMARY` はリブランド後も一貫して「主色」を指し、参照側
（slides.py / スキル / サンプル）の import を変えずに済む。名前が色相を約束しない
ので、主色を青や緑に変えても `PRIMARY` のまま嘘にならない。
既定値は暖色オレンジ基調（`PRIMARY` = #EC6739）。

リブランド手順（詳細は docs/REBRAND.md）:
  1. 下の色・フォント・TEMPLATE を自社ブランドに書き換える
  2. `templates/` に自社雛形 .pptx を置き、TEMPLATE をそのファイル名にする
     （雛形は現雛形と同じレイアウト構成・プレースホルダを満たすこと）
"""
from pptx.dml.color import RGBColor

# ============================================================ ブランドカラー
# 役割ベースの定数名。値（RGBColor）だけを書き換え、名前は変えない（参照側が依存）。
# `add_icon()` の color= のみ hex 文字列、それ以外は RGBColor。
PRIMARY       = RGBColor(0xEC, 0x67, 0x39)   # 主色 / 見出しアクセント / 推奨側
PRIMARY_LIGHT = RGBColor(0xFB, 0xE2, 0xD6)   # 淡い主色 / 差し色
SECONDARY     = RGBColor(0x1F, 0x4E, 0x79)   # 補色 / 信頼・構造
SUCCESS       = RGBColor(0x5E, 0x8B, 0x7E)   # 成功・正常
DANGER        = RGBColor(0xC2, 0x54, 0x50)   # 警告・否定
HIGHLIGHT     = RGBColor(0xB0, 0x85, 0x1B)   # 強調・推奨バッジ
TEXT_MUTED    = RGBColor(0x6B, 0x6B, 0x6B)   # 中立・サブテキスト
TEXT          = RGBColor(0x43, 0x43, 0x43)   # 本文テキスト
BORDER        = RGBColor(0xD5, 0xD5, 0xD5)   # 罫線
SURFACE       = RGBColor(0xF5, 0xF5, 0xF5)   # カード背景
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)   # 反転テキスト・明色背景（絶対色）
BLACK         = RGBColor(0x00, 0x00, 0x00)   # 絶対色

# ============================================================ フォント
# 和文 JP_FONT は配布先に無いとフォールバックする（README / CLAUDE.md の落とし穴参照）。
# latin を変える場合、検査ゲートの幅計測がそのフォントを参照する点に注意。
FONT    = "Arial"          # latin（英数字）
JP_FONT = "Noto Sans JP"   # 和文（ea/cs）
JOSEFIN = "Josefin Sans"   # 表紙・章番号など装飾英字

# ============================================================ 既定雛形
# templates/ 配下の雛形 .pptx のファイル名。自社雛形に差し替えたらここを変更する。
TEMPLATE = "スライド雛形.pptx"
