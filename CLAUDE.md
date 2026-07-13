# スライド自動生成ガイド

このディレクトリは、ブランドフォーマットに沿った **スライド（PowerPoint）を自動生成する作業用** です。
雛形 `templates/スライド雛形.pptx` をベースに `python-pptx` で各種スライドを作成します。

## 雛形仕様

- **スライドサイズ・フォント**: 25.40 × 14.29 cm（16:9）／ Arial（英数字 latin）＋ Noto Sans JP（和文 ea/cs）。英字アクセントは Josefin Sans。`_set_run` が latin=Arial・和文=`JP_FONT` を明示割当し、Arial の無制御 CJK フォールバックを防ぐ。タイプスケール定数 `T_H1`=24 / `T_H2`=15 / `T_BODY`=10.5 / `T_CAPTION`=8.25 (pt)。 **色・フォント・既定雛形ファイル名は `brand.py` に一元定義**し、`slides.py` が再エクスポートする（リブランドは `brand.py` を編集）。
- **レイアウトゾーン**: 見出し `y = 0.0–3.0 cm`／コンテンツ `y = 3.0–12.8 cm`、横 `x = 1.07–24.33 cm`。見出し領域には自分で要素を描かない（テンプレが描く）。

### レイアウト一覧（`slides.py` の定数で参照可能）

| 定数 | idx | レイアウト名 | 用途 |
|---|---|---|---|
| `L_BACK` | 0 | 裏紙 | 裏表紙 |
| `L_CHAPTER` | 1 | TITLE_AND_BODY | 章扉（全面主色） |
| `L_BODY` | 2 | ページ番号+タイトル社外秘 | 本文（主色枠＋ページ番号＋CONFIDENTIAL） |
| `L_COVER` | 3 | TITLE | 表紙（ロゴ＋タイトルテキスト） |

### Layout 2（本文）のプレースホルダ

| プレースホルダ | 位置 (cm) | サイズ (cm) | 用途 |
|---|---|---|---|
| section label | y=0.94 | 13.67 × 0.68 | "OVERVIEW / 全体コンセプト" 等の小ラベル |
| title | y=1.42 | 22.41 × 1.41 | ページ大タイトル |
| page# | (23.77, 0.46) | 1.16 × 1.06 | 右上の番号サークル（主色円） |

## slides.py のヘルパー早見表

```python
from slides import (
    PRIMARY, PRIMARY_LIGHT, SECONDARY, SUCCESS, DANGER, HIGHLIGHT, TEXT_MUTED, TEXT, BORDER, SURFACE, WHITE, JOSEFIN, JP_FONT,
    T_H1, T_H2, T_BODY, T_CAPTION,
    L_BODY, L_CHAPTER, SLIDE_W_CM,
    textbox, shape_box, card, callout, connector, picture,
    bar_chart, line_chart, pie_chart, CHART_PALETTE,
    configure_body, configure_chapter, update_cover,
    load_template, reset_to_cover_only, add_back_cover, finalize_page_numbers,
    add_icon, validate,
    ICON_TREE, ICON_CERT, ICON_EYE, ICON_KEY, ICON_LOCK, ICON_FINGERPRINT,
    ICON_PHONE, ICON_USER_SHIELD, ICON_CHECK, ICON_X, ICON_ARROW,
)
```

### ビルダー
- `textbox(slide, x, y, w, h, text, *, size, bold, color, align, jp_font, ...)` — 任意のテキストボックス
- `shape_box(slide, MSO_SHAPE.X, x, y, w, h, *, text, fill, line, jp_font, shadow=False, ...)` — 任意の図形＋テキスト。**影は既定オフ（フラット標準）**、`shadow=True` で従来のカード影を opt-in。
- `connector(slide, x1, y1, x2, y2, *, color, width, dash, begin_arrow, end_arrow)` — 直線/コネクタ（概念図・フロー図でボックス同士をつなぐ）。色は `RGBColor`（既定 `TEXT_MUTED`）
- `picture(slide, image, x, y, w=None, h=None)` — 画像（写真・ロゴ・図版）。`w`/`h` の片方省略でアスペクト比保持。存在しないパスは `FileNotFoundError`

### フラットコンポーネント（本文の基本単位）
- `card(slide, x, y, w, h, *, title, body, chip=None, kicker=None, accent=PRIMARY, ...)` — 枠なし白カード（ヘアライン枠＋影なし）。アクセントチップ＋濃色テキスト見出し＋本文。色ベタの見出し帯は使わない。
- `callout(slide, x, y, w, h, text, *, accent=PRIMARY, ...)` — 結論・キーテイクアウェイ用の淡色ピル＋左アクセントバー（反転白文字の色帯を置き換える）。

### グラフ（ブランド配色）
- `bar_chart(slide, x, y, w, h, *, categories, series, colors=None, horizontal=False, show_value=True, number_format=None, ...)` — 棒グラフ。`series` は `[("系列名", [値,...]), ...]`
- `line_chart(slide, x, y, w, h, *, categories, series, colors=None, markers=True, ...)` — 折れ線グラフ
- `pie_chart(slide, x, y, w, h, *, categories, values, colors=None, show_value=True, show_percentage=False, ...)` — 円グラフ
- 既定パレットは `CHART_PALETTE`（`PRIMARY/SECONDARY/SUCCESS/HIGHLIGHT/TEXT_MUTED/DANGER`）。`colors=` で系列/扇ごとに上書き可

### アイコン挿入
- `add_icon(slide, icon_name, x, y, size=1.0)` — Iconify から取得した PNG を cm 座標で配置（取得後は `~/.cache/ksr-slides/icons/` にキャッシュ）

### スライドタイプ設定
- `configure_body(slide, *, section_label, title)` — Layout 2 のプレースホルダ設定
- `configure_chapter(slide, *, chapter_num, title, subtitle)` — Layout 1 章扉設定
- `update_cover(slide, *, lines)` — サンプル表紙のテキスト書き換え（lines は dict のリスト）

### ワークフロー
- `load_template()` — `templates/スライド雛形.pptx` を読み込み（検知レジストリもリセット）
- `reset_to_cover_only(prs)` — サンプルの章扉/本文/裏紙を削除（表紙のみ残す）
- `add_back_cover(prs)` — 裏表紙スライドを末尾に追加
- `finalize_page_numbers(prs, skip_first=True)` — 全ページにページ番号を挿入

### レンダリング（5-b 視覚確認用）
- `render_pngs(pptx_path, *, dpi=120)` — pptx を全ページ PNG 化して保存パスのリストを返す。`validate()` が作った PDF をそのまま再利用し（mtime 比較で同期）、PyMuPDF で PNG 化する（soffice の二重起動なし）。

### レイアウト検査（はみ出し・見切れ・ゾーン逸脱の機械検知）
- `validate(prs, OUT, render=True)` — `generate.py` 末尾で呼ぶ必須ゲート。生成時リント（`validate_fit`）＋レンダ後突合（`validate_render`）を実行しレポート表示。**ERROR が 1 件でもあれば exit code 1 で終了する**（既定 `strict=True`。レポートのみで続行したい場合だけ `strict=False`）。
- 各ビルダー（`textbox`/`shape_box`/`configure_*`/`update_cover`）が描いた要素を自動でレジストリ記録するため、追加実装は不要（`validate` を呼ぶだけ）。
- 非テキスト要素（`connector`/`picture`/`bar_chart`/`line_chart`/`pie_chart`）も geom レジストリに配置矩形を記録し、本文コンテンツゾーン逸脱を `ZONE` ERROR で検知する。ただし**グラフ内部のテキスト（軸ラベル・データラベル・凡例）は文字照合（`validate_render`）の対象外**＝配置（ゾーン）は検査されるが、内部テキストの見切れは自動検知されないので 5-b 視覚確認で見る。
- 依存: `Pillow`・`pymupdf`（`pip install pymupdf` 済み前提）。`pymupdf` が無い場合はレンダ後突合だけスキップし生成時リントは実行。

## テンプレ固定デザイン（再設計禁止・上書き禁止）

表表紙・裏表紙・セクション扉・見出し、およびテンプレートから引き継がれる装飾（右上ページ番号サークル／CONFIDENTIAL バッジ／本文の白背景・装飾主色バー・罫線／表紙・裏表紙のロゴ／扉の全面主色）は、現状のデザインをそのまま使う。下記ワークフロー（テンプレ起点で構築）に従えばレイアウト継承で自動保持される。自分で描き直したり図形で隠したりしない。AI が差し替えるのは**テキスト内容のみ**：

| 部分 | 実装 | AI が触れる範囲 |
|---|---|---|
| 表表紙 | `update_cover(prs.slides[0], lines=[...])` | テキストのみ（色は常に `WHITE`） |
| セクション扉 | `add_slide(L_CHAPTER)` ＋ `configure_chapter(...)` | 章番号・タイトル・サブタイトルのテキストのみ |
| 見出し | `add_slide(L_BODY)` ＋ `configure_body(section_label, title)` | セクションラベル・タイトルのテキストのみ |
| ページ番号 | `finalize_page_numbers(prs, skip_first=True)` | 呼ぶだけ |
| 裏表紙 | `add_back_cover(prs)` | 呼ぶだけ |

## アイコンは任意

アイコンは [Iconify API](https://iconify.design) 経由で `add_icon(slide, "mdi:アイコン名", x, y, size, color="#EC6739")` の形で任意に挿入できる（事前ファイル不要 / 検索: https://icon-sets.iconify.design/）。`slides` には `ICON_TREE` `ICON_CERT` `ICON_KEY` `ICON_LOCK` `ICON_CHECK` `ICON_X` 等の定数も用意されている。**挿入するかどうか・どこに置くかは内容から自由に判断**してよい（必須ではない）。`add_icon()` の `color=` のみ hex 文字列、それ以外の色引数は `RGBColor`。

---

## 既知の落とし穴と回避策

### ⚠ はみ出し・見切れは目視に頼らず検査ゲートで検知する
テキストの枠はみ出し・末尾見切れ・意図しない改行は実行時エラーにならず、PNG 目視でも非決定的にしか拾えない（特に `shape_box` の溢れ文字は LibreOffice が描画ごと消すため目視で気付けない）。`generate.py` 末尾で必ず `validate(prs, OUT, render=True)` を呼び、`❌ ERROR` を解消するまで完了としない（ERROR 時は exit code 1 で失敗するため、シェル・CI からも機械的に検知できる）。`validate_fit` は生成時近似（行高 ≒ `size×1.30×line_spacing` で計測。PowerPoint＋CJK フォールバックの上振れを保守側に近似した係数）、`validate_render` は実レンダラ準拠の確定検知：描画文字列が意図より短い＝`CLIP`、**1文字も描画されていない完全消失＝`LOST`**、レンダ PDF のページ欠落＝`PAGE_MISSING`、ページゾーン逸脱＝`ZONE`、**個別 `shape_box` の親矩形からの実測 bbox 逸脱＝`SHAPE_OVERFLOW`**。さらに `validate_fit` は、低い塗り矩形（色帯ヘッダ・バッジ等、高さ ≤ `BAND_MAX_H`=3cm）に載せた textbox がその帯の下端を突き抜ける場合を `BAND_OVERFLOW`＝**ERROR** に格上げする（色帯の外に文字が漏れるのは確定的な崩れ。背の高い白カードは帯とみなさず従来どおり WARN）。検知は SLACK 0.07cm/相対4%で誤検知を抑制済み。ビルダー以外で直接 `add_textbox`/`add_shape` した要素はレジストリに乗らず検知対象外になるため、テキスト要素はヘルパー経由で描く。

### ⚠ PowerPoint 実描画は LibreOffice より行送りが大きい（セーフマージン必須）
LibreOffice ベースの `validate_render` が PASS しても、PowerPoint で開くと Arial＋CJK フォールバック（MS P ゴシック等）の行送り差で固定高 `shape_box` を破ることがある。`validate_fit` の係数 1.30 で部分的に近似しているが、設計側でも **`need_h ≤ 0.85 × box_h`（または上下合計 0.30cm 以上）の余白**を必ず取る。`need_h` と `box_h` が 0.1cm 未満で接していたら PASS でも設計やり直し。レイヤー帯・手順カード・テーブル行のように行数の多い固定高シェイプで特に要注意。

### ⚠ 和文 Noto Sans JP は配布先に無いとフォールバックする（埋め込み or PDF で配布）
本文・見出しの和文は `JP_FONT="Noto Sans JP"`。Noto Sans JP は **Windows に標準搭載されていない**ため、未インストールの PC で `.pptx` を開くと別の和文ゴシックにフォールバックし、行送り・字形が崩れる（python-pptx はフォント埋め込み非対応）。顧客へ**編集可能な .pptx で渡す場合**は (1) 先方に Noto Sans JP を導入してもらう、(2) PowerPoint/LibreOffice で開き直してフォント埋め込み保存する、(3) **配布は PDF 化する**、のいずれかで字形を固定する。社内 mac での確認・PDF 納品なら導入済みでそのまま可。別フォントに切り替えたい場合は `JP_FONT` を変更（latin の Arial は据え置きで `validate` の幅計測は不変）。
