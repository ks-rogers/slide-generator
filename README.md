# スライド自動生成キット

ブランドフォーマットに沿ったスライド (.pptx) を **Python + Claude Code** で自動生成するキット。提案資料はもちろん、勉強会・社内資料・登壇資料など各種スライドに使える。雛形 PowerPoint と再利用可能なライブラリ (`slides.py`)、そして 2 つの Claude Code スキル (`/slides-new`・`/slides-gen`) を組み合わせ、資料ごとに 10〜20 枚のスライドを一気に出力する。

---

## 全体フロー

```
┌──────────────────┐    ┌──────────────────┐    ┌──────────────────┐
│ /slides-new      │ -> │ /slides-gen      │ -> │ output.pptx      │
│ 壁打ちで原稿作成    │    │ generate.py 生成 │    │ 検査 PASS + 視覚   │
│ -> content.txt   │    │ -> 実行 -> 検査   │    │   確認済み         │
└──────────────────┘    └──────────────────┘    └──────────────────┘
```

各ステップは Claude Code 内のスラッシュコマンドで起動する。手動で `generate.py` を書く必要はない（微調整時のみ後から触る）。

---

## 0. 事前準備（初回のみ）

> **想定環境：** 本キットは **macOS / Windows / Linux** で動作する（開発・動作確認は macOS 中心）。§0-2 の表に **macOS（Homebrew）と Windows（winget／公式インストーラ）** の導入例を併記した。Linux は各ディストリのパッケージマネージャ（`apt` / `dnf` 等）に読み替える。

### 0-0. リポジトリを取得する

```bash
git clone https://github.com/ks-rogers/slide-generator.git
cd slide-generator
```

clone して出来るフォルダ名は `slide-generator/`。以降の手順はすべて**このフォルダの直下（プロジェクトルート）**で実行する。

### 0-1. Python 依存（venv で隔離する）

> **手で打ちたくない場合：** Claude Code 上で `/slides-setup` を実行すると、この 0-1・0-2 の作業（venv 作成・依存インストール・システム依存チェック・検証）をスキルが代行する（`git clone`〔0-0〕と Claude Code 本体の導入〔0-3〕は対象外）。下記は手動で行う場合の手順。

ホスト Python を汚さないよう **プロジェクト直下に venv を作って**そこにインストールする。以降のすべての `python3` 呼び出しはこの venv のものを使う。

```bash
# プロジェクトルート（slide-generator/）で
python3 -m venv .venv              # 初回のみ
source .venv/bin/activate          # 新しいシェルを開くたび
pip install -r requirements.txt    # 初回のみ
```

`(.venv)` がプロンプトに付いていれば有効化済み。`deactivate` で抜ける／`rm -rf .venv` で完全削除できる。

`requirements.txt` の中身（参考）：

| パッケージ | 用途 |
|---|---|
| `python-pptx` | pptx 出力本体 |
| `PyMuPDF` (fitz) | `validate_render` と `render_pngs` のレンダラ（**必須**） |
| `Pillow` | `validate_fit` の幅計測 |
| `requests` / `cairosvg` | `add_icon` で Iconify から SVG 取得・PNG 変換（**アイコン使用時のみ**） |

検査ゲートの回帰テスト（`tests/`）を回す場合のみ、開発用依存を追加する（本番生成には不要）。`requirements-dev.txt` は本番依存を含まないため、本番依存とあわせて指定する：

```bash
pip install -r requirements.txt -r requirements-dev.txt   # 本番依存 + pytest
pytest                                                     # tests/ を実行（main への PR 時は CI でも自動実行）
```

### 0-2. システム依存（pip 外・ホスト側に必要）

venv では隔離できない OS レベルの依存。

> **必須の 3 つ：Python・git・LibreOffice**（この 3 つが無いと `/slides-gen` が完走しない）。cairo はアイコンを使うときだけ、パッケージマネージャは導入を楽にするための任意ツール。

| ツール | 必須度 | 用途 | macOS インストール | Windows インストール |
|---|---|---|---|---|
| **Python 3.10 以上** | **必須** | 生成の実行環境そのもの（venv／`python-pptx` で pptx を組み立てる。`slides.py` が PEP 604 記法・`Pillow>=12.3` を使用） | `brew install python@3.12`（または `python@3.11` など） | [python.org 公式インストーラ](https://www.python.org/downloads/windows/)（**「Add python.exe to PATH」に必ずチェック**）／または `winget install --id Python.Python.3.12` |
| **git** | **必須** | リポジトリ取得／`/slides-gen`・`/slides-setup` がプロジェクトルート解決に `git rev-parse` を使用 | Xcode CLT 同梱（`xcode-select --install`）／または `brew install git` | [Git for Windows](https://git-scm.com/downloads/win)／または `winget install --id Git.Git` |
| **LibreOffice** (`soffice`) | **必須** | `validate_render`（pptx → PDF）／`render_pngs` 前段。検査ゲートの実体 | `brew install --cask libreoffice` | [公式インストーラ](https://www.libreoffice.org/download/download/)／または `winget install --id TheDocumentFoundation.LibreOffice`（導入後、`soffice` の PATH 通しが必要な場合あり → 下記） |
| **cairo** | 任意（アイコン時のみ） | `cairosvg` のランタイム依存（`add_icon` 使用時のみ） | `brew install cairo libffi` | 標準では同梱されず用意がやや手間。**アイコンを使わない資料ならスキップ可**。入れる場合の手順は下記「cairo（アイコン使用時のみ）の入れ方」参照 |
| **パッケージマネージャ** | 任意（推奨） | 上記の導入をコマンドで一括化する前提。無くても各公式インストーラで代替可 | [Homebrew](https://brew.sh)（`brew install` の前提） | **winget**（Windows 10/11 に標準同梱。`winget install` の前提）。[Chocolatey](https://chocolatey.org) 等でも可 |

> Python 依存では `PyMuPDF`（fitz）も**必須**（§0-1 の `requirements.txt` に含む。pip で入る）。上表の LibreOffice(`soffice`) と合わせて、この 2 つが検査ゲート（手順 5）の実体。

必須の Python・git・LibreOffice、および `PyMuPDF` が揃っていることを確認しておく（揃っていないと **`/slides-gen` の検査ゲートが動かず手順 5 で停止する**）。venv 有効化後、次の 1 行で **Python 版・`soffice` のパス・PyMuPDF の読み込み**の 3 点をまとめて確認できる：

```bash
# macOS / Linux
python3 --version && which soffice && python3 -c "import fitz; print(fitz.__doc__[:40])"
```

```powershell
# Windows (PowerShell)。venv 有効化は .venv\Scripts\Activate.ps1、python は python（python3 ではない）
python --version; Get-Command soffice; python -c "import fitz; print(fitz.__doc__[:40])"
```

> **`which soffice` が何も表示されないとき：** `brew install --cask libreoffice`（cask 版）は `soffice` を PATH に登録せず、実体を `/Applications/LibreOffice.app/Contents/MacOS/soffice` に置く。次のいずれかで PATH を通す（`<bin>` は `/opt/homebrew/bin` 等、PATH 内の書き込み可能なディレクトリ）：
>
> ```bash
> # 方法A: シンボリックリンクを張る
> ln -s /Applications/LibreOffice.app/Contents/MacOS/soffice <bin>/soffice
> # 方法B: シェルの設定ファイル（~/.zshrc 等）に PATH を追加
> export PATH="/Applications/LibreOffice.app/Contents/MacOS:$PATH"
> ```
>
> 設定後に `which soffice` がパスを返せば OK。

> **Windows で `soffice` が見つからないとき：** `slides.py` は PATH 上の `soffice` を探す（`shutil.which("soffice")`）。Windows 版 LibreOffice の実体は `C:\Program Files\LibreOffice\program\`（CLI 用の `soffice.com` と GUI 用の `soffice.exe` を同梱）にあり、既定では PATH に登録されないため、このフォルダを PATH に追加する：
>
> ```powershell
> # 一時的に通す（そのターミナル/セッション内だけ有効）
> $env:Path += ";C:\Program Files\LibreOffice\program"
> ```
>
> 恒久的に通すなら「設定 → システム → バージョン情報 → システムの詳細設定 → 環境変数」で、ユーザーの `Path` に `C:\Program Files\LibreOffice\program` を追加する（追加後は Claude Code／ターミナルを開き直す）。設定後、`Get-Command soffice` がパスを返せば OK。

> **ネットワーク：** `add_icon()` でアイコンを使う場合、初回はランタイムに [Iconify API](https://api.iconify.design) から SVG を取得する（取得後は `~/.cache/ksr-slides/icons/` にキャッシュされ、以降オフラインで利用可）。アイコンを使わない資料であれば不要。

> **cairo（アイコン使用時のみ）の入れ方：** `add_icon()` を使わないなら不要（`requirements.txt` の `cairosvg`/`requests` を入れても、実行時に呼ばなければ問題ない）。使う場合、`cairosvg` は OS 側の cairo（＋ libffi）を必要とする（[CairoSVG 公式](https://cairosvg.org/documentation/)の案内に準拠）：
>
> - **macOS**：`brew install cairo libffi`
> - **Linux**：`sudo apt install libcairo2 libffi-dev python3-dev`（Debian/Ubuntu 系）／`sudo dnf install cairo libffi-devel python3-devel`（Fedora 系）
> - **Windows**：cairo は標準で入らないため、[GTK for Windows Runtime Environment Installer](https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer) で cairo（`libcairo-2.dll`）を導入する（CairoSVG 公式が案内する方法。ビルドを伴う環境では [Visual C++ Build Tools](https://visualstudio.microsoft.com/visual-cpp-build-tools/) も要ることがある）。
>
> OS 側を入れた後に `pip install -r requirements.txt`（または `/slides-setup`）で `cairosvg`/`requests` を導入する。導入確認は `python -c "import cairosvg; print('cairosvg OK')"`。

### 0-3. Claude Code

`/slides-new` `/slides-gen` は Claude Code のスキル機能で動くため、Claude Code 本体のインストールが必要。共通の前提として **Anthropic アカウントでのサインイン（または API キー設定）** が済んでいること。

スキルが認識される条件は **CLI・デスクトップアプリ共通で「本リポジトリのルート（`slide-generator/`）が作業ディレクトリになっていること」**（`.claude/skills/slides-new` `slides-gen` をそこから読む）。使い方に応じて以下のどちらかで起動する。

#### A. CLI で使う場合（macOS / Linux / Windows）

- インストール：[公式ドキュメント](https://docs.claude.com/en/docs/claude-code/overview) の手順に従う（npm 経由：`npm install -g @anthropic-ai/claude-code`、または各種ネイティブインストーラ）
- 0-0 で clone したフォルダに `cd` してから `claude` を起動する：

  ```bash
  cd slide-generator
  claude
  ```

#### B. デスクトップアプリで使う場合（macOS / Windows）

> デスクトップアプリは Linux 非対応。Linux では A の CLI を使う。

1. [公式ダウンロードページ](https://code.claude.com/docs/en/desktop)から macOS / Windows 版をインストールし、起動してサインイン後、**Code タブ**を開く（Windows は初回のみ [Git for Windows](https://git-scm.com/downloads/win) が必要。インストール後アプリを再起動）。
2. **New session**（サイドバーの「+ New session」／ macOS `Cmd+N`・Windows `Ctrl+N`）で新しいセッションを作り、**Project folder に 0-0 で clone した `slide-generator/` フォルダを選択**する。これでそのフォルダが作業ディレクトリになり、スキルが認識される。
3. 0-1・0-2 の venv 構築・依存インストール（`python3 -m venv .venv` 〜 `pip install ...`）は、**統合ターミナル**（**Views** メニュー、または macOS / Windows とも ``Ctrl+` ``）で実行する。統合ターミナルはセッションの作業ディレクトリで開き、Claude と同じ環境を共有するため、CLI と同じ手順がそのまま使える。

> **補足（Git リポジトリのセッション分離）：** デスクトップアプリは Git リポジトリのセッションごとに [git worktree](https://code.claude.com/docs/en/worktrees) で隔離コピーを作る（既定で `<project-root>/.claude/worktrees/`）。`.venv/` は `.gitignore` 済みのため worktree 側には複製されない。**各セッションの統合ターミナルで、そのセッションの作業ディレクトリに対して 0-1・0-2 を実行する**こと（セッションをまたいで `.venv` は共有されない）。

---

## 1. 新規スライド資料を作る（`/slides-new`）

Claude Code で次を実行する：

```
/slides-new
```

引数としてタイトル・テーマを渡してもよい（省略すると壁打ちで聞かれる）：

```
/slides-new サンプル銀行 AIチャットボット導入
```

### このスキルがやること

まず **既存の原稿（下書き・ドラフト・資料メモ）があるかを確認**し、フローが分岐する：

- **原稿あり** → その原稿（チャット貼り付け or ファイルパス）を読み込んで構成にマッピングし、`content.txt` 形式に整形する。原稿から読み取れない項目だけをヒアリングで補完する。
- **原稿なし** → 以下の壁打ち（ゼロからのヒアリング）で作成する。

いずれの場合も最終成果物は `/slides-gen` が読み込める `content.txt`。

1. **ヒアリング**（原稿で埋まらない部分のみ）
   - タイトル・テーマ・背景課題・核となるメッセージ・ターゲット読者・枚数目安 等
   - 不明な数値・固有名詞は `【要確認】〇〇` プレースホルダで残す（推測でねつ造しない）
2. **構成案の提示と合意**
   - 「P.1 表紙 / P.2 全体コンセプト / PART 01 … / 裏表紙」の章構成を提案し、ユーザーが OK したら次へ
3. **ページごとの内容ヒアリング**
4. **`projects/<資料フォルダ>/content.txt` を生成**
   - フォルダ名は `<資料名>`（提案資料なら `<顧客名略称>_<テーマ略称>`。記号・スペース不可）。例：`社内勉強会_LLM入門` / `サンプル銀行_AIチャットボット`
   - フォーマットは `/slides-gen` が解析できる **80 文字 `=` セパレータ＋`【タグ】` 形式**（後述）

### 完了時の状態

```
projects/<資料フォルダ>/
└── content.txt       ← 原稿（要確認箇所は【要確認】で残っている）
```

> **次に進む前に：** `content.txt` を開き、`【要確認】` プレースホルダを必ず実情報で埋める。ここで埋めずに `/slides-gen` に進むと、確認漏れがそのままスライドに乗る。

---

## 2. content.txt から PPTX を生成する（`/slides-gen`）

Claude Code で次を実行する：

```
/slides-gen projects/<資料フォルダ>/content.txt
```

引数を省略するとカレントディレクトリの `content.txt` を探索する。**資料フォルダに `cd` してから引数なしで叩く**のも可：

```
cd projects/<資料フォルダ>
/slides-gen
```

### このスキルがやること

1. **検証環境チェック**（`soffice` / `PyMuPDF`）— 不足があれば停止して指示を仰ぐ
2. **参照ファイルの読み込み**（`content.txt` / `slides.py` / `design-guide.md`）
3. **ページ設計**（content.txt をパース → ページ種別判定 → 各ページの構成を設計）
4. **`generate.py` を生成**（テンプレ起点：`load_template` → `update_cover` → `reset_to_cover_only` → 各ページ → `add_back_cover` → `finalize_page_numbers` → `validate`）
5. **実行 → レイアウト検査ゲート (`validate`)**
   - 生成時リント (`validate_fit`)：折返し近似で縦あふれ・横はみ出し・ゾーン逸脱を検知。色帯（高さ ≤ 3cm の塗り矩形）からテキストが漏れる場合は `BAND_OVERFLOW` (ERROR) として検知
   - レンダ後突合 (`validate_render`)：soffice→PDF→PyMuPDF で実描画を計測し、`CLIP` / `ZONE` / `SHAPE_OVERFLOW` / `WRAP` を検知
   - `❌ ERROR` が出たら**完了不可**（`validate` は ERROR があると exit code 1 で終了する）。座標・サイズ・テキスト量を調整して PASS するまで再生成
6. **視覚確認ゲート**（`render_pngs` で全スライドを PNG 化 → `Read()` で目視）
   - design-guide.md §1〜9 の準拠監査を表形式で書き出し、1 項目でも No なら再設計
7. **完了報告**（`output.pptx` のパス＋監査表）

### 完了時の状態

```
projects/<資料フォルダ>/
├── content.txt
├── generate.py       ← /slides-gen が生成
└── output.pptx       ← 検査 PASS + 視覚確認済み
```

---

## 3. 出力の確認と微調整

`output.pptx` を **PowerPoint と Google スライドの両方で開いて目視確認**する。

軽微な微修正が必要なときは：

- **テキスト内容の差し替えだけ** → `content.txt` を編集し再度 `/slides-gen`
- **特定ページのレイアウト調整** → `generate.py` を直接編集し、資料フォルダで `python3 generate.py` を再実行（venv 有効化済みであること。末尾の `validate(prs, OUT, render=True)` が再走する）

---

## ディレクトリ構成

> `.claude/tmp/ksr-slides/` は clone 直後には存在せず、検査・視覚確認時の PNG/PDF 生成で自動生成される。`add_icon()` のアイコン PNG はリポジトリ外の `~/.cache/ksr-slides/icons/`（ユーザースコープ。資料・worktree 間で共有）にキャッシュされる。

```
slide-generator/
├── README.md                         ← 本ファイル（人間向け）
├── LICENSE                           ← MIT ライセンス
├── CONTRIBUTING.md                   ← コントリビューションガイド
├── CLAUDE.md                         ← AI エージェント向けの作業ガイド
├── requirements.txt                  ← Python 依存（システム依存はコメント参照）
├── requirements-dev.txt              ← 開発用依存（pytest／本番生成には不要）
├── pyproject.toml                    ← pytest 設定（testpaths / pythonpath）
├── slides.py                       ← 共通ヘルパーライブラリ
├── brand.py                        ← ブランド設定の単一ソース（色・フォント・雛形名／リブランドはここ）
├── templates/
│   └── スライド雛形.pptx           ← マスター雛形（編集禁止推奨）
├── projects/                         ← 資料ごとのフォルダ（中身は git 管理外／.gitkeep のみ追跡）
│   └── <資料名>/                     ← 以下はすべてローカル限定（コミットされない）
│       ├── content.txt               ← 原稿（/slides-new が生成）
│       ├── generate.py               ← ページ定義（/slides-gen が生成）
│       └── output.pptx               ← 生成結果
├── examples/                         ← 公開サンプル資料（架空・git 追跡）
│   ├── README.md
│   └── サンプル商事_AIチャットボット/  ← content.txt / generate.py / output.pptx / preview.png
├── docs/
│   └── REBRAND.md                    ← 自社ブランド向けに作り変える手順
├── tests/                            ← pytest（検査ゲートの回帰テスト）
│   ├── conftest.py
│   └── test_*.py                     ← BAND_OVERFLOW / CONTAINER_OVERFLOW / ZONE 等の回帰
├── .github/
│   └── workflows/tests.yml           ← CI（main への PR 時に pytest 実行）
└── .claude/
    ├── skills/
    │   ├── slides-setup/SKILL.md     ← /slides-setup スキル定義（環境構築）
    │   ├── slides-new/SKILL.md       ← /slides-new スキル定義
    │   └── slides-gen/
    │       ├── SKILL.md              ← /slides-gen スキル定義
    │       └── design-guide.md       ← 設計品質バー（slides-gen が遵守）
    └── tmp/ksr-slides/               ← (自動生成) 中間 PNG / PDF
```

---

## content.txt のフォーマット

`/slides-new` が `projects/<資料フォルダ>/content.txt` を自動で書き出す。80 文字 `=` セパレータでページを区切り、各ページは `【タグ】` 形式でフィールドを並べる構造。ページ種別（表紙／セクション扉／本文／裏表紙）はヘッダ `P.X ／ ...` の文字列で判定する。

- **フォーマット詳細仕様**（テンプレ、各種タグ、セクションラベル英語部分の例）→ [`.claude/skills/slides-new/SKILL.md`](.claude/skills/slides-new/SKILL.md) §4
- **ページ種別の判定ルール（実装側）** → [`.claude/skills/slides-gen/SKILL.md`](.claude/skills/slides-gen/SKILL.md) §3

本文ページの構成（カード / 2 列比較 / フロー / 表 / KPI / 階層図 …）は `/slides-gen` が content.txt の情報の型から **その都度設計**する（固定テンプレなし）。

---

## 雛形仕様・レイアウトゾーン

技術仕様（スライドサイズ 25.40 × 14.29 cm／フォント Arial／レイアウト一覧／Layout 2 のプレースホルダ／コンテンツゾーン `y=3.0〜12.8 cm`・`x=1.07〜24.33 cm` など）は [`CLAUDE.md`](CLAUDE.md) §雛形仕様 を参照。`validate` ゲートがゾーン逸脱を ERROR にするので、コンテンツは必ずこの矩形内に収める。

---

## ブランド設定（`brand.py` で一元管理）

ブランドカラー・フォント・既定雛形ファイル名は [`brand.py`](brand.py) に**一元定義**している（単一ソース）。**リブランドは `brand.py` を編集するだけ**（手順は [docs/REBRAND.md](docs/REBRAND.md)）。コードからは `slides` 経由で定数としてインポートし（`brand.py` を `slides.py` が再エクスポート）、**色はこの定数以外を使わない**（`/slides-gen` の遵守事項）。HEX 値の実体は `brand.py` のコメントを参照（README には重複させない）。

| 定数 | 想定用途 |
|---|---|
| `PRIMARY` / `PRIMARY_LIGHT` | ブランド主色／薄い差し色 |
| `SECONDARY` | 構造・補色／信頼系 |
| `SUCCESS` / `DANGER` / `HIGHLIGHT` | 成功・正常／警告・否定／強調・推奨 |
| `TEXT_MUTED` / `TEXT` / `BORDER` / `SURFACE` | 中立サブ／本文テキスト／罫線／背景 |
| `WHITE` | 反転テキスト・明色背景 |

フォント（`FONT`=latin／`JP_FONT`=和文／`JOSEFIN`=装飾英字）・既定雛形（`TEMPLATE`）も同じく `brand.py`。`python-pptx` の色引数は `RGBColor(r, g, b)` 型（hex 文字列は不可）。例外として `add_icon()` の `color=` のみ hex 文字列 (`"#EC6739"`)。

---

## トラブルシューティング

### セットアップ・実行時のエラー

| 症状 | 対処 |
|---|---|
| `/slides-gen` が「`soffice` が無い」で停止する | `brew install --cask libreoffice` を実行し、`which soffice` でパスを確認 |
| `/slides-gen` が「`.venv/` が無い」で停止する | プロジェクトルートで `python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt` |
| `import fitz` で失敗 | venv 有効化を確認（`which python3` が `.venv/bin/python3` を指すか）。未インストールなら `pip install -r requirements.txt` |
| `pip install` が `externally-managed-environment` で拒否される | venv 未有効化。`source .venv/bin/activate` してから再実行 |
| `add_icon` で `cairo` エラー | `brew install cairo`（cairo は OS 側必須。`cairosvg`/`requests` は venv で `pip install -r requirements.txt` 済みのはず） |
| `validate` が `❌ ERROR` を吐く | 該当ページの座標・サイズ・テキスト量を `generate.py` で調整して再実行。content.txt の情報は自己判断で削らない |
| ページ番号が出ない | `finalize_page_numbers(prs, skip_first=True)` が main 末尾で呼ばれているか確認 |

### `generate.py` を手で触るときの注意

`/slides-gen` は以下を踏まえて生成するため通常は問題ないが、微調整で直接編集する場合は要注意：

- **PowerPoint と LibreOffice の行送り差**：固定高 `shape_box` で `validate` PASS でも PowerPoint で枠を破ることがある既知の落とし穴。設計時の対応ルール（`need_h ≤ 0.85 × box_h` 等）は [`CLAUDE.md`](CLAUDE.md) §既知の落とし穴 を参照。
- **`\n` は pptx で改行にならない**：複数行は `["1 行目", "2 行目"]` のリストで渡す。
- **日本語＋数字／記号混在テキストが □（tofu）になる（LibreOffice 表示のみ）**：`shape_box` のテキストを `[文字列]` のリストにし、フォントサイズを整数 pt に。`validate` では検知不能なので視覚確認で目視する。
- **表紙を `add_slide(L_COVER)` で作らない**：ロゴ画像が消える。テンプレ既存スライド（idx 0）を `update_cover()` で書き換える方式が必須。
- **グラフは `bar_chart` / `line_chart` / `pie_chart` で対応**（ブランド配色・凡例・データラベル付き）。⚠ グラフ内部テキスト（軸・ラベル・凡例）は `validate` の文字照合対象外なので、潰れ・はみ出しは視覚確認で見る。配置（コンテンツゾーン）は自動検査される。
- **接続線・図版・写真は `connector` / `picture` で対応**：概念図の線/矢印は `connector(...)`、画像配置は `picture(...)`（片側省略でアスペクト比保持）。どちらも配置矩形がゾーン検査される。
- **アニメーション・複雑な SmartArt は未対応**：`python-pptx` の制約。必要なら手動で追加。

---

## リファレンス資料

完成例として [`examples/サンプル商事_AIチャットボット/`](examples/) に**架空資料のサンプル**（`content.txt` → `generate.py` → `output.pptx`／検査 PASS 済み）を同梱している。入力と成果物の対応を見たいときはここから。

実際の `projects/` 配下は各自のローカル作業用で **git 管理対象外**（機密情報を含むため `.gitignore` 済み。フォルダ自体は `projects/.gitkeep` で維持される）。手元に既存資料があれば `generate.py` の実装例として参照できるが、**レイアウト構成・デザインパターンをそのまま転用しない**（新しい資料は内容から自由に設計し直す）。

---

## 自社ブランドで使う（リブランド）

同梱の既定雛形は K.S.ロジャースのブランドに沿っているが、**仕組み（ヘルパーと検査ゲート）はブランド非依存**で、雛形 `.pptx`・配色・フォントを差し替えれば自社向けに転用できる（エンジンのコード変更は不要）。手順は [`docs/REBRAND.md`](docs/REBRAND.md) を参照。

---

## ライセンス

コードは [MIT License](LICENSE)（Copyright © 2026 K.S.ロジャース株式会社）。コントリビューションは [CONTRIBUTING.md](CONTRIBUTING.md) を参照。

> ⚠️ `templates/スライド雛形.pptx` に含まれる **K.S.ロジャースのロゴ・ブランドデザインは同社の資産**であり、MIT（コード）の対象ではない。自社・他者が利用する場合は上記「リブランド」に従って雛形ごと自社ブランドに差し替えること（KSR ロゴを載せたまま配布しない）。

> **埋め込みフォント：** `templates/スライド雛形.pptx` および `examples/サンプル商事_AIチャットボット/output.pptx` には、装飾英字の **Josefin Sans**（Regular / Bold / Italic / Bold Italic の4ウェイト・サブセット）がフォント実データとして埋め込まれている。Josefin Sans は **SIL Open Font License 1.1**（`Copyright 2010 The Josefin Sans Project Authors`、Reserved Font Name "Josefin Sans"）。著作権表示とライセンス全文は [`NOTICE`](NOTICE) を参照。なお **Arial・Noto Sans JP は名前参照のみで非同梱**（埋め込みは Josefin Sans のみ）。

### サードパーティ依存のライセンス

本リポジトリのコードは MIT だが、実行時に import する依存は各自が `pip install` するもので、それぞれのライセンスに従う（本リポジトリには同梱していない）。特に以下は **コピーレフト系** なので、本キットを組み込んだ製品を配布・SaaS 提供する場合は各ライセンスの義務を確認すること。

| パッケージ | ライセンス | 用途 | 備考 |
|---|---|---|---|
| `python-pptx` | MIT | pptx 出力本体 | 寛容 |
| `Pillow` | MIT-CMU (HPND) | `validate_fit` の幅計測 | 寛容 |
| `requests` | Apache-2.0 | `add_icon`（アイコン使用時のみ） | 寛容 |
| `lxml` / `XlsxWriter` | BSD | python-pptx の推移依存 | 寛容 |
| **`PyMuPDF`** | **AGPL-3.0 / 商用デュアル**（Artifex） | `validate_render`・`render_pngs` のレンダラ | **コピーレフト**。派生物の配布・ネットワーク提供時は AGPL 義務が及びうる。回避したい場合は Artifex の商用ライセンスを取得するか、レンダリング（検査ステップ 5-a/5-b）を使わない運用にする |
| **`cairosvg`** | **LGPL-3.0+** | `add_icon` の SVG→PNG 変換（アイコン使用時のみ） | **弱コピーレフト**。動的リンク（別プロセス／別ライブラリとしての利用）なら本体は MIT のまま |

- レンダリング（PyMuPDF）とアイコン（cairosvg/requests）は **どちらもオプション** で、未インストールでもコア機能（生成・生成時リント）は動作する（該当ステップを degrade してスキップ）。AGPL/LGPL を避けたい配布形態ではこれらを除外できる。
- `add_icon` が [Iconify](https://iconify.design) から取得するアイコンは、**アイコンセットごとのライセンス**（MIT / Apache / CC 等）に従う。生成物に載せて配布する際は各アイコンの出所ライセンスを確認すること。
