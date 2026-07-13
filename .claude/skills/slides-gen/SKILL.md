---
name: slides-gen
description: content.txtを読み込んでgenerate.pyを自動生成する
argument-hint: "[content.txtのパス（省略時はカレントディレクトリを探索）]"
---

# スライド generate.py 自動生成スキル

`content.txt` の内容を解析し、`slides.py` を使った `generate.py` を生成・実行する。

## このスキルの基本方針

**コンテンツは自由に設計してよい。遵守すべき制約は次の2つだけ**（詳細は CLAUDE.md・README.md 参照）：

1. **ブランドカラー**：`slides` の色定数のみを使う（`PRIMARY` `PRIMARY_LIGHT` `SECONDARY` `SUCCESS` `DANGER` `HIGHLIGHT` `TEXT_MUTED` `TEXT` `BORDER` `SURFACE` `WHITE`）。`textbox()` / `shape_box()` / `fill=` は `RGBColor(r,g,b)` 型、`add_icon()` の `color=` のみ hex 文字列。
2. **レイアウトゾーン**：本文の見出し領域 `y=0.0〜3.0` には自分で描かない（テンプレが描く）。コンテンツは `y=3.0〜12.8 / x=1.07〜24.33` の矩形内に収める。

それ以外（構図・図形・配色の組み合わせ方・アイコンの有無・ヘルパー関数を使うかどうか）は、ページごとに「どう見せれば一番わかりやすく美しいか」で**完全に自由**に決める。固定パターンや必須テクニックは存在しない。

## テンプレ固定デザインは再設計しない・上書きしない

表表紙・裏表紙・セクション扉・見出し、およびテンプレが描く装飾（右上ページ番号サークル／CONFIDENTIAL／本文の白背景・主色バー・罫線／表紙裏表紙ロゴ／扉の全面主色）は、下記ワークフロー（テンプレ起点で構築）に従えば**レイアウト継承で自動保持される**。自分で描き直したり上に図形を重ねて隠したりしない。AI が差し替えるのは**テキスト内容のみ**（表紙の御中・タイトル、扉の章番号・タイトル、見出しのセクションラベル・タイトル）。実装関数の対応表は CLAUDE.md「テンプレ固定デザイン」を参照。

## 前提条件

- カレントディレクトリが本プロジェクト内であること
- `slides.py` がプロジェクトルート（`projects/` の2階層上）にあること
- **Python 依存はプロジェクトルート直下の venv（`.venv/`）に隔離してインストール済みであること**。ホスト Python は使わない。セットアップ手順は `requirements.txt` 冒頭コメントまたは [README §0-1](../../../README.md#0-1-python-依存venv-で隔離する) 参照。
- **以降のすべての Python 呼び出しは `.venv/bin/python3`（プロジェクトルートからの相対パス）を使う**。Claude Code の Bash ツールはコマンド間でシェル状態が永続しないため `source .venv/bin/activate` に頼らず、`.venv/bin/python3` を明示的に呼ぶ（または `"$(git rev-parse --show-toplevel)/.venv/bin/python3"` で絶対パス化）。
- **LibreOffice（`soffice`）・`PyMuPDF`(fitz) が利用可能であること（必須）**。手順5の検査ゲート（`validate_render`）と視覚確認（`render_pngs`）はこの2つに依存し、代替アプリ（PowerPoint / Keynote）は自動ゲートの代替にしない。未導入時は手順1で停止する（下記）。

## 手順

### 1. ファイルの特定

引数が指定された場合はそのパスを content.txt として使用する。指定がなければ次の順で探索：

1. カレントディレクトリの `content.txt`
2. IDE で開いているファイルが content.txt であればそのパス

`generate.py` の出力先は `content.txt` と同じディレクトリ。

**続行前に検証環境を確認する**（生成してから手順5でコケて「検証できないまま完了扱い」になるのを防ぐ）。プロジェクトルートで実行：

```bash
which soffice \
  && [ -x .venv/bin/python3 ] \
  && .venv/bin/python3 -c "import fitz"
```

どれかが失敗したら**生成に進まず停止し**、状況に応じて以下をユーザーに伝えて指示を仰ぐ：

- `soffice` が無い → `brew install --cask libreoffice`
- `.venv/bin/python3` が無い → `python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt`
- `import fitz` が失敗 → venv 有効化下で `pip install -r requirements.txt`

盲目生成して `output.pptx` だけ残し「完了」と報告しない。`requests` / `cairosvg` はアイコン使用時のみ必要なので、ここでは検査しない。

### 2. 参照ファイルの読み込み

並列で読み込む：
- `content.txt`（生成対象）
- プロジェクトルートの `slides.py`（利用可能な関数・定数の確認）
- このスキルの `design-guide.md`（**拘束力あり**。任意の参考資料ではない）

**`design-guide.md` の位置づけ：** §1〜§9 と §8 の6問チェックは、全本文ページの設計（手順3・3.5・4）と監査（5-b）に**必ず**適用する。「コンテンツは自由に設計してよい」の自由は design-guide の枠内での自由であり、design-guide 違反は遵守事項1・2（色／ゾーン）違反と同格の**完了不可事由**として扱う。CLAUDE.md の「アイコンは任意」は *アイコンという手段* が任意という意味であって、design-guide §9 の「全スライドに視覚要素（図・チャート・アイコン・色面・構造図形のいずれか）を1つ以上」は拘束力がある（テキストだけのスライドを作らない）。「アイコン任意」を口実に視覚要素ゼロ・箇条書き流し込みのページを正当化しない。

### 3. content.txt の解析とページ設計

`================== P.X ／ ...` のヘッダでページを分割し、各ページを分類：

| 判定条件 | スライドタイプ | 構築方法 |
|---|---|---|
| `表紙` を含む | 表表紙 | `update_cover(prs.slides[0], ...)` |
| `セクション扉` を含む、または `PART XX` がある | セクション扉 | `add_slide(L_CHAPTER)` ＋ `configure_chapter(...)` |
| `裏表紙` を含む | 裏表紙 | `add_back_cover(prs)`（main で呼ぶ） |
| その他 | 本文 | `add_slide(L_BODY)` ＋ `configure_body(...)` ＋ 自由コンテンツ |

本文ページは、内容を読んで「読み手が最初に理解すべきことは何か」「テキストを並べるより図・色・余白・サイズ差で表現できないか」を考え、ページごとに最適な構成を**自由に**設計する。カード／結論ピルは **`card()` / `callout()`（フラット標準）を基本単位**にし、それ以外は `shape_box()` ＋ `textbox()` ＋ `add_icon()` を組み合わせて自由にレイアウトする（色ベタ見出し帯・影は使わない＝design-guide §7）。記号（✓✗→等）をアイコンに変換するかどうかも内容次第で自由に判断する。

**content.txt の情報は欠落させない**（忠実性の要件であってデザイン制約ではない）。レイアウトに収まらないと判断しても自己判断で内容を削らず、必要ならユーザーに確認する。

### 3.5 設計品質バー（手抜き設計の禁止）

**`validate` ゲートは「はみ出し/見切れ/ゾーン逸脱」しか見ない。ゲート通過＝良い設計ではない。** 「自由に設計してよい」は「楽をしてよい」という意味ではなく、「内容に最適な構造を毎ページ考える」という意味。次のいずれかに該当するページは**手抜きとみなし、完了不可。再設計する**。

**禁止アンチパターン（該当したら必ず作り直す）:**
- 縦長ボックスに短いテキストを中央寄せで置き、上下に大きな空白が空く（＝コンテンツゾーン `y=3.0〜12.8` を使い切っていない）
- 種類の違う情報を1つの箇条書きに流し込んだだけ（ラベルと値・比較・手順などの構造を表現していない）
- 全角スペースやコロンで桁揃えを偽装している（→ セル分割した表で整列させる）
- 情報の型が違うのに全ページ同じ「ヘッダ帯＋箇条書き」になっている
- 余白が「意図した余白」ではなく「埋めなかった結果」になっている
- **角丸カード（`ROUNDED_RECTANGLE`）の上に直角矩形（`RECTANGLE`）を重ねている**：カード上端のヘッダー帯、左端の番号ブロック、内側の塗り分けなどを `RECTANGLE` で乗せると、外側カードの R と内側の直角がコーナーで噛み合わず「屋根が浮いた」ダサい見た目になる。ヘッダー帯や帯状の塗り分けが必要な構成では、**外側カード自体も `RECTANGLE`（直角）にする**（角丸×直角を入れ子にしない）。角丸を使うのは「単独の浮いたカード」や pill／円のように上に何も乗せない箱だけに限定する。**結論バー（`callout()`）の淡色下地も同じ理由で直角**：角丸ピルの左端に直角のアクセントバーを重ねると、左端の上下でバーと下地の間に三日月状の隙間が出て崩れる（`slides.callout()` は直角で実装済み。自前で結論バーを組むときも下地＝直角矩形にする）。
- **テーブルの外枠・行・セルに角丸を使う**：表全体を `card()`（`ROUNDED_RECTANGLE`）で囲ったり、各行を角丸でくくったりしない。表は罫線つき `RECTANGLE` で直角に組む。表は「整列・密度・規則性」が読みやすさの本体であり、角丸はそれを"ふやけさせる"。
- **見出し・タイトル・サブタイトルの下に自前の装飾アンダーライン（短い主色細バー等）を引いている**：区切り・強調などの機能がない飾りの線は "AI 生成の手癖" の典型（design-guide §9）。引いてよい線は「機能がある」もの（表の罫線、ラベル行の区切り下線など）だけ。テンプレが描く見出しの主色バーは固定デザインなので残す（自分で足さない・消さない）。装飾の細バー用ヘルパー（`accent_bar` 的なもの）を作らない。`validate` は検知しないので 5-b 監査で必ず潰す。

**情報の型 → 既定で優先検討する構造（強制ではないが、安易な箇条書きより必ず先に検討する）:**

| content の型 | 推奨構造 |
|---|---|
| 属性と値の一覧（ポリシー／仕様）| ラベル列＋値列の2カラム表（セル分割で整列） |
| 2者以上の比較（従来 vs 提案 等）| 横並びカラム（`card()`：チップ＋濃色見出し） |
| 手順・フロー | ステップカード＋矢印（横／縦） |
| 多項目の対応表（役割分担 等）| グリッド表（カテゴリ見出し行で区切る） |
| 概念の構成・階層 | 図解（ボックス＋接続線）。箇条書きにしない |
| 少数の重要指標 | 大きな数値タイル |

各本文ページは設計時に必ず「この内容の型は何か」「コンテンツゾーン全体を使った構造になっているか」「自分が客に出す資料としてこれを見て雑だと思わないか」を自問する。**具体の作り方（フォント ladder・余白・整列・色運用・図解化の判断基準）は `design-guide.md` に従う**（本節が"禁止"、design-guide が"こう作る"）。

**コードを書く前に、各本文ページについて以下を確定させる（design-guide 由来・過去に頻発した実害）:**

1. **死に余白の予防（§3）**：縦長カード/パネル/タイルに少量テキストを入れる設計を**先に**潰す。コンテンツが短いなら ①箱を内容高に合わせて縮める、②指標タイル・アイコン・図で意味のある要素を足してゾーンを満たす、のどちらかを設計時に決める。「とりあえず大きい箱に箇条書き＋MIDDLE寄せ」は禁止。`line_spacing` を 1.8 以上に広げて空白を埋めるのは"埋めなかった結果の余白"であり禁止（行間で誤魔化さない）。
2. **多カラム本文の折返し（§5）**：狭い列に小さいフォントで長文を入れると単語/数値が行中で割れる（例「明文化」→「明／文化」、「99.5%以／上」）。`validate` はこれを ERROR にしない。列幅・文言長・フォントサイズを設計時に見積もり、トークンが割れない幅／長さにする。最小 8pt は床だが、8pt でも割れるなら列を広げるか文言を縮める。
3. **視覚要素とモチーフ（§9）**：各本文ページに視覚要素を1つ以上持たせる方針を決め、資料全体で反復するモチーフを1つ選ぶ。情報の型が同じページは同じ見せ方（反復）、型が違うページだけ構造を変える。「型が違うのに全ページ同じ帯＋箇条書き」も「型が同じなのにページごとにバラバラ」も禁止。
4. **固定高 `shape_box` の縦方向セーフマージン（PowerPoint 互換）**：複数行テキストを入れる固定高シェイプ（`shape_box(text=..., fill≠None)`／カードの内側／表セル／レイヤー枠など、内容を**クリップする**枠）は、設計時点で **`need_h ≤ 0.85 × box_h`**（または **上下合計 0.30cm 以上の余白**）を確保する。**LibreOffice で `validate` PASS していても、`need_h` と `box_h` が 0.1cm 未満で接していたら設計やり直し**（ルールの背景・行送り係数の理屈は [`CLAUDE.md`](../../../CLAUDE.md) §既知の落とし穴 §⚠ PowerPoint 実描画 を参照）。レイヤー帯・横並びの行（手順カード等）では行数増で誤差が累積するので、行数 ×（フォントサイズの 2 倍 / 28.35）cm ぶんは最低の上下マージンとして取る。

### 4. generate.py の生成

**必須の骨組み**（テンプレ起点で構築 → 固定デザインが自動保持される）：

```python
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

from slides import (
    PRIMARY, PRIMARY_LIGHT, SECONDARY, SUCCESS, DANGER, HIGHLIGHT,
    TEXT_MUTED, TEXT, BORDER, SURFACE, WHITE, JOSEFIN, JP_FONT,
    T_H1, T_H2, T_BODY, T_CAPTION,
    L_BODY, L_CHAPTER, SLIDE_W_CM,
    textbox, shape_box, card, callout, connector, picture,
    bar_chart, line_chart, pie_chart, CHART_PALETTE,
    configure_body, configure_chapter, update_cover,
    load_template, reset_to_cover_only, add_back_cover, finalize_page_numbers,
    add_icon, validate,
)

OUT = Path(__file__).parent / "output.pptx"

# 各ページを slide_p1, slide_p2, ... として定義
# 本文ページの先頭は必ず:
#   s = prs.slides.add_slide(prs.slide_layouts[L_BODY])
#   configure_body(s, section_label="...", title="...")
# その後 y=3.0〜12.8 の範囲だけに自由にコンテンツを描く

def main():
    prs = load_template()
    slide_p1(prs.slides[0])          # 表紙はテンプレの既存スライドを書き換える
    reset_to_cover_only(prs)         # サンプルの扉/本文/裏紙を削除

    # slide_p2(prs) ... slide_pN(prs) を順番に呼ぶ

    add_back_cover(prs)
    finalize_page_numbers(prs, skip_first=True)
    prs.save(OUT)
    print(f"saved: {OUT}")
    print(f"slides: {len(prs.slides)}")
    # レイアウト検査（生成時リント＋レンダ後突合）。ERROR は必ず解消する。
    validate(prs, OUT, render=True)

if __name__ == "__main__":
    main()
```

### 5. 確認・修正ループ（必須）

「生成 → 検査 → 視覚確認 → 問題があれば修正して4に戻る」を回す。完了条件は 5-c。

---

#### 5-a. 実行・レイアウト検査（第一ゲート）

```bash
cd <content.txtのディレクトリ> && "$(git rev-parse --show-toplevel)/.venv/bin/python3" generate.py
```

（venv の python3 を絶対パスで呼ぶ。プロジェクトルートから `.venv/bin/python3 projects/<資料名>/generate.py` でも可だが、`cd` してから絶対パス指定の方が `generate.py` 内の相対パス参照と整合する）

実行時エラーは原因を調査して修正・再実行する。エラーがなくても、`generate.py` 末尾の `validate(prs, OUT, render=True)` が**レイアウト検査レポート**を出力する。ERROR が 1 件でもあれば `validate` は **exit code 1** で終了する（＝generate.py の実行自体が失敗として見える。既定 `strict=True`）。

`validate` が行う2段階の検知：

- **生成時リント**（`validate_fit`）— PIL+Arial で折返しを近似計測し、縦あふれ・横はみ出し・コンテンツゾーン逸脱を検出。行高は `size × 1.30 × line_spacing`（PowerPoint＋CJK フォールバックを近似する保守係数）で見積る。横はみ出しのスラックは `shape_box` で `w < 3cm` のとき相対 0%・絶対 0.02cm に厳格化（pill/badge の hairline 折返しを通常スラックが見逃すのを防ぐ）。
- **レンダ後突合**（`validate_render`）— soffice→PDF を PyMuPDF で実測し、意図テキストと照合。描画文字列が意図より短い＝見切れ確定（`CLIP`）、**1 文字も描画されていない完全消失（`LOST`）**、レンダ PDF のページ欠落（`PAGE_MISSING`）、実測 bbox のページゾーン逸脱（`ZONE`）、**個別 `shape_box` の親矩形からの実測 bbox 逸脱（`SHAPE_OVERFLOW`）**、**1段落想定の固定枠（高さが 1 行分しかない pill/badge/タイルラベル等）が実描画で 2 行以上に折返した状態（`WRAP`）** を検出。

| 結果 | 対応 |
|---|---|
| `✅ PASS` | 次の視覚確認（5-b）へ進む。 |
| `❌ ERROR`（`OVERFLOW_V`/`OVERFLOW_H`/`ZONE`/`CLIP`/`SHAPE_OVERFLOW`/`WRAP`/`BAND_OVERFLOW`） | **完了不可。** 座標・サイズ・テキスト量を調整して generate.py を修正し、PASS するまで 5-a を繰り返す（content.txt の情報は削らない／手順3参照）。 |
| `⚠️ WARN` | 固定枠の文字量過多・textbox の縦伸長など。5-b の視覚確認で実害を確認し、崩れていれば ERROR と同様に修正。**`BAND_OVERFLOW`（色帯ヘッダ等の塗り帯から textbox がはみ出す）は WARN ではなく ERROR で確定検知される**（帯高を上げる／フォントを下げる／配置を見直す）。 |

---

#### 5-b. 視覚確認（第二ゲート）

はみ出し・見切れ・ゾーン逸脱は 5-a で機械検知済みなので、ここでは**機械が拾えない美観・設計品質**をチェックする。**これは単なる崩れ確認ではなく設計批評ゲート**であり、ここで手抜きを発見したら ERROR 同等（完了不可）として 3.5 の品質バーに従い再設計する。中間ファイルは `.claude/tmp/ksr-slides/` に書く：

```bash
"$(git rev-parse --show-toplevel)/.venv/bin/python3" -c "
import sys; from pathlib import Path
sys.path.insert(0, str(Path('<output.pptxのパス>').resolve().parents[2]))
from slides import render_pngs
pngs = render_pngs('<output.pptxのパス>')
print(f'{len(pngs)} slides at {pngs[0].parent}')
"
```

`render_pngs()` は 5-a の `validate()` が作った PDF を再利用するので、soffice の二重起動はしない（mtime チェックで pptx より新しければスキップ。pptx を再生成すれば自動で再変換）。PNG レンダリングは PyMuPDF が担当する（pdf2image / poppler 不要）。

全スライド画像を `Read()` で確認する。

**■ design-guide.md 準拠監査（必須・書面・スキップ不可）**

前提（design-guide §8 / anthropics QA 哲学）：**「問題は必ずある。それを見つけるのが仕事」**。指摘ゼロは"よく見ていない"だけ。最低1回は「指摘→修正→再確認」を通すまで完了不可。

各本文ページについて、以下を**応答本文に表形式で書き出す**（頭の中で済ませない／「全ページOK」の一括宣言は不可）。1ページ1行で、§8 の6問と下記3つの実害チェックの結果・根拠を記す：

| 観点 | 判定基準（No なら完了不可＝再設計して 5-a へ） |
|---|---|
| §8-1 1メッセージ | 結論が1つに絞れ、見出し／サブで伝わるか |
| §8-2 階層 | サイズ→太さ→色の順、強調1〜2点に収まるか |
| §8-3 余白の意図 | **死に余白チェック**：カード/パネル/タイル内に >1.0cm の意図しない空白がないか。縦長ボックスに短文＋中央寄せの"浮き"がないか。`line_spacing≥1.8` で空白を誤魔化していないか |
| §8-4 整列 | セル分割で整列しているか（全角空白／コロン桁揃えの偽装でないか）|
| §8-5 型に合う構造 | 3.5 の「型→推奨構造」に従い、箇条書き流し込みでないか／図解すべきを図解したか |
| §8-6 色の意味 | ブランド色のみ・色が意味に対応・均等配分でなくドミナンス階層（§9）か |
| 折返し（§5） | **多カラム本文で単語/数値/トークンが行中で割れていないか**（`validate` は拾わない。画像で必ず目視）|
| 視覚要素（§9） | 各本文ページに視覚要素が1つ以上あるか／全体で1モチーフを反復し、型が同じページの見せ方が揃っているか |
| 装飾線（§9） | **見出し/タイトル/サブタイトルの下に機能のない装飾アンダーライン（短い主色細バー等）を引いていないか**（AI手癖。`validate` は拾わない。引いてよいのは表罫線・ラベル区切り等の機能ある線だけ。テンプレ描画の主色バーは残す） |

**1ページでも1項目でも No／該当があれば、ERROR と同格の完了不可。** 3.5 の品質バーと design-guide に従い generate.py を再設計し、5-a から再実行する。

その他のチェック観点：
- **【tofu】日本語＋数字・記号混在テキストが □（tofu）になっていないか**。`validate` はコードポイント一致で判定するため tofu を検知できず、ここでしか拾えない（対処はトラブルシューティング参照）
- WARN 指摘箇所が実際に崩れていないか
- 要素同士が意図せず重なっていないか（検知層が拾わない色違い図形の重なり等）
- 使用色がブランドカラー表の範囲内か
- 余白・サイズ差・図形バランスが読みやすいか（余白は「意図した余白」か）
- テンプレ固定デザイン（表紙ロゴ／扉の全面主色／見出しの主色番号サークル＋ラベル＋タイトル＋白背景＋装飾バー＋CONFIDENTIAL／全ページのページ番号／裏表紙ロゴ）が崩れていない・隠れていないか

---

#### 5-c. 判定

- **問題なし** → 5-b の design-guide 準拠監査表を全本文ページ分書き出し、**全行・全項目が No／該当なし**であることを確認できて初めて完了。`output.pptx` のパスと、監査表をユーザーに伝える。
- **問題あり** → generate.py を修正して **5-a から再実行**する。監査表に1つでも No／該当があれば「問題あり」。

## 技術的注意（デザインの良し悪しではなく「正しく動かす」ための制約）

- **`\n` は pptx で改行にならない**。複数行テキストはリストで渡す。NG: `"1行目\n2行目"` / OK: `["1行目", "2行目"]`。`shape_box()` の text もリスト要素に `\n` を含めない。
- **`textbox()` は縦方向に auto-expand する**。高さ上限がある場所では `shape_box()`（透明背景 `fill=None`、無枠 `line=None`）を使う。`shape_box()` はクリップし `anchor=MSO_ANCHOR.MIDDLE` が正確に効く。
- **色引数に hex 文字列は不可**。`textbox()` / `shape_box()` / `fill=` は `RGBColor(r, g, b)` で渡す（`from pptx.dml.color import RGBColor`）。`add_icon()` の `color=` のみ hex 文字列。
- **フォント代替で行が溢れる前提で安全マージンを取る**（具体的な数値は 3.5 §4、症状別の対処は末尾トラブルシューティング参照）。検証は LibreOffice のみで PowerPoint・Google スライドの実描画はこのフローでは確認できないため、固定高 `shape_box()` の文字量をギリギリに詰めず、オートフィット（縮小して収める）にも依存しない。
- **寸法は cm 単位**。位置計算は `SLIDE_W_CM = 25.40` を基準に。`Pt()` はフォントサイズのみ。
- **ページ番号は自動描画されない**。`add_slide(layout)` は `sldNum` を継承しないため、`finalize_page_numbers(prs)` を main の最後で必ず呼ぶ。
- **表紙を `add_slide(L_COVER)` で作らない**。ロゴ画像が消える。テンプレの既存スライド（idx 0）を `update_cover()` で書き換える方式が必須（`load_template` → `update_cover(prs.slides[0], ...)` → `reset_to_cover_only(prs)` → 以降を `add_slide`）。
- **`configure_body` の section_label / title はテンプレ固定**でフォント・サイズ・位置を変更できない（変更しようとしない）。コンテンツ領域 `y=3.0` 以下のみ自由。

## トラブルシューティング

| 症状 | 原因 | 対処 |
|---|---|---|
| ページ番号が出ない | `finalize_page_numbers` 未呼び出し | main の最後で呼ぶ |
| 表紙のロゴが消える | `add_slide(L_COVER)` で新規作成した | `update_cover` で書き換える方式に |
| 本文ページに枠や CONFIDENTIAL が出ない | 間違ったレイアウト指定 | `prs.slide_layouts[L_BODY]` を使う |
| テキストがはみ出す／見切れる | 文字量過多・固定高 shape_box の溢れ | サイズ・座標・テキスト量を調整 |
| 章扉のサブタイトルが大きすぎる | `subtitle_size` 既定値 | `configure_chapter(..., subtitle_size=11)` |
| 日本語＋数字／`・`混在の文字が □（tofu）になる | LibreOffice の script itemization バグ。run の rPr は正常（latin=Arial／ea・cs=Noto Sans JP）・`validate_render` も PASS で **pptx 自体は正常**（PowerPoint/Google では正常表示）。LibreOffice が digit/`・`↔漢字境界で非CJKフォントを誤選択する LO 固有の描画バグ。同一文字列でも単一文字列渡しは tofu・リスト渡しは正常になる | 該当 `shape_box()`／`textbox()` のテキストを単一文字列でなく **`[文字列]` のリスト形式**で渡し、フォントサイズを**整数pt**にする（例 `text=v` → `text=[v]`, `size=9.5` → `size=10`）。`validate` では検知不能なので 5-b 視覚確認で必ず目視する |
| ページが間延び・箇条書き流し込みで雑（手抜き設計）| 固定パターンへの逃げ（3.5 設計品質バー違反）。`validate` は破綻しか見ないため通過してしまう | 3.5 の「情報の型→推奨構造」で再設計し、コンテンツゾーンを使い切る。5-a PASS でも 5-b でこれを必ず批評する |
| **LibreOffice では収まったが PowerPoint で枠を破る（行送り溢れ）** | PowerPoint の Arial＋CJK フォールバックは LibreOffice より行送りが大きい既知の落とし穴（詳細は [CLAUDE.md §既知の落とし穴 §⚠ PowerPoint 実描画](../../../CLAUDE.md)） | 設計時に **`need_h ≤ 0.85 × box_h`** または上下合計 0.30cm 以上の余白を確保（3.5 設計品質バー §4）。`validate_render` の `SHAPE_OVERFLOW` は LibreOffice 実測でも親矩形からの bbox 逸脱を ERROR にするが、検査 PASS でも 5-b で「枠端ギリギリ」を見たら設計余裕を疑う |
