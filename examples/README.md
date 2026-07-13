# サンプル資料

`/slides-gen` がどんな入力からどんな成果物を作るかを示すリファレンス実装。
**ここに含まれる顧客名・数値はすべて架空**で、実在の企業・資料とは関係ありません。

> 資料フォルダ本体は `projects/`（`.gitignore` 済み・ローカル限定）に作りますが、
> 公開リポジトリにはサンプルが無いと挙動が伝わらないため、`examples/` に1件だけ
> サニタイズ済みの完成例を同梱しています。

## サンプル商事_AIチャットボット

社内問い合わせ対応 AIチャットボット導入提案（全8ページ）。

| ファイル | 内容 |
|---|---|
| `content.txt` | 原稿（`/slides-new` が作る形式。ページ構成の入力） |
| `generate.py` | ページ定義（`/slides-gen` が作る形式。`slides` のヘルパーで組み立て） |
| `output.pptx` | 生成結果（検査ゲート `validate` PASS 済み） |
| `preview.png` | 全ページのプレビュー一覧 |

![プレビュー](サンプル商事_AIチャットボット/preview.png)

### 自分で生成してみる

リポジトリルートで venv を有効化し（[README](../README.md) の「0. 事前準備」）、
次を実行すると `output.pptx` が再生成されます。

```bash
source .venv/bin/activate
python3 examples/サンプル商事_AIチャットボット/generate.py
```

末尾の `validate(prs, OUT, render=True)` がレイアウト検査を実行し、
`✅ PASS` まで通ることを確認できます。

### デモンストレーションしている要素

- 表紙の書き換え（`update_cover`）／セクション扉（`configure_chapter`）／本文見出し（`configure_body`）
- フラットな白カード（`card`）とキーテイクアウェイのコールアウト（`callout`）
- 2カラム比較（AS-IS / TO-BE）と、意味に対応した配色（`DANGER` / `SUCCESS`）
- ステップフロー（`card` ＋ `connector` の矢印）
- ブランド配色の棒グラフ（`bar_chart`）
- 検査ゲート（`validate`）を通る座標・サイズ設計
