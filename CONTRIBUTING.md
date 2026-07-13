# コントリビューションガイド

**バグ報告・改善提案は [Issue](../../issues/new/choose) で歓迎します。** テンプレートに
沿って起票してください（機密情報は貼らないこと）。

> ℹ️ 本リポジトリは**利用（MIT）と閲覧を目的とした公開**で、**外部からの Pull Request
> （fork 経由）は受け付けていません**（fork を無効化しています）。コードの変更は
> K.S.ロジャース組織メンバーが下記の手順で行います。改善アイデアがあれば Issue で
> お知らせください。以降の「開発環境」「設計上の約束」「コミット・PR」は**組織メンバー
> 向け**の手順です。

## 開発環境のセットアップ

[README](README.md) の「0. 事前準備」に従って venv とシステム依存（LibreOffice 等）を
用意してください。テストを回すには開発用依存も入れます。

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt -r requirements-dev.txt
pytest
```

`pytest`（`tests/` の検査ゲート回帰テスト）は **main への PR で CI でも自動実行**されます
（[.github/workflows/tests.yml](.github/workflows/tests.yml)）。ローカルで PASS してから
PR を出してください。

## 設計上の約束

- **検査ゲートを通すこと**：スライドを生成するコードは `generate.py` 末尾で
  `validate(prs, OUT, render=True)` を呼び、`❌ ERROR` を解消してから完了とします
  （詳細は [CLAUDE.md](CLAUDE.md) の検査ゲート節）。
- **テンプレ固定デザインを上書きしない**：表紙・裏表紙・章扉・見出しや、テンプレートが
  描く装飾（主色バー・ページ番号サークル・CONFIDENTIAL バッジ等）は継承で保持します。
  図形で隠したり描き直したりしないでください。
- **テキスト要素はヘルパー経由で**：`textbox` / `shape_box` / `card` / `callout` などを
  使います。素の `add_textbox` / `add_shape` は検査レジストリに乗らず検知対象外になります。
- **ハードコードを避ける／エラーを握り潰さない**：色やサイズはブランド定数・タイプスケール
  定数を使い、例外は黙って捨てないでください。
- 自社ブランド向けの転用は [docs/REBRAND.md](docs/REBRAND.md) を参照（エンジンは変更不要）。

## コミット・PR

- コミットメッセージは [Conventional Commits](https://www.conventionalcommits.org/) 形式、
  本文は日本語可（例：`feat: 棒グラフのデータラベル書式を追加`）。
- 1 PR = 1 つの関心事に絞ると review しやすくなります。
- 機密情報（実在顧客名・実案件の `content.txt` / `generate.py` / `output.pptx` 等）を
  含めないでください。案件フォルダ `projects/` は `.gitignore` 済みです。
