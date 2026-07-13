---
name: slides-setup
description: 事前準備（venv 作成・依存インストール・システム依存チェック）を自動実行し、CLI を手で叩かずに環境構築を完了させる
argument-hint: "[--with-tests（pytest 用の開発依存も入れる場合）]"
---

# スライド 環境セットアップスキル

README §0（事前準備）の作業を肩代わりするスキル。ユーザーが `python -m venv` や `pip install` などを手で打たなくても、このスキルがプロジェクトルートを解決し、依存をチェック・インストールして検証まで通す。**macOS / Linux / Windows に対応**（Windows では Claude Code の Bash ツールが使う Git Bash 上で動く前提。README §0-3 のとおり Git for Windows が必要）。

## このスキルの基本方針

- **値を捏造しない／盲目的に進めない**：不足しているシステム依存（`soffice` 等）を勝手に「入っていることにして」先に進めない。検出 → 提示 → 必要なら導入、の順を守る。
- **システムを変更する操作（`brew install` / `winget install` 等）は実行前にユーザーへ確認する**。Python 依存（venv・pip）はプロジェクト内に閉じるので確認なしで進めてよい。
- **`source .venv/bin/activate` に頼らない**：Claude Code の Bash ツールはコマンド間でシェル状態が永続しないため、venv の Python（後述の `PYBIN`）を**明示的に**呼ぶ。
- **OS 差を吸収する**：venv の Python パスやパッケージ管理コマンドは OS で異なる（§1 の対応表）。以降のコマンドは `PYBIN`（venv の Python）を毎回自動判定して使うので、mac / Linux / Windows(Git Bash) で共通に動く。
- **対象は §0-1・0-2 の環境構築のみ**。`git clone`（§0-0）は本スキル起動前にリポジトリが手元にある前提（スキルはリポジトリ内で動くため）。Claude Code 本体のインストール（§0-3）も対象外。

---

## 手順

### 1. プロジェクトルートの解決・OS 判定・前提確認

```bash
ROOT="$(git rev-parse --show-toplevel)" && echo "ROOT=$ROOT"
ls "$ROOT/requirements.txt" "$ROOT/slides.py"
case "$(uname -s)" in
  Darwin) OSKIND=mac;;
  Linux)  OSKIND=linux;;                 # WSL もここ（apt/dnf 等）
  MINGW*|MSYS*|CYGWIN*) OSKIND=windows;;  # Git Bash
  *) OSKIND=unknown;;
esac
echo "OSKIND=$OSKIND"
```

- `git rev-parse` が失敗する／`requirements.txt`・`slides.py` が見当たらない場合は、**カレントディレクトリが本リポジトリ（`slide-generator/`）の中にない**可能性が高い。停止して「本リポジトリのルートで起動し直す（CLI なら `cd slide-generator`、デスクトップアプリなら Project folder に当該フォルダを選択）」よう案内する。
- 以降のコマンドはすべて `"$ROOT"` を基準に絶対パスで実行する（`cd` してもよいが、Bash 間で状態は持ち越されないので毎回 `"$ROOT/..."` を使うのが安全）。

**OS 別の対応表**（以降のステップで参照）：

| OS (`OSKIND`) | venv の Python (`PYBIN`) | パッケージ管理 | `soffice` の既定パス（PATH 未登録時の在処） |
|---|---|---|---|
| `mac` | `.venv/bin/python3` | Homebrew（`brew`） | `/Applications/LibreOffice.app/Contents/MacOS/soffice` |
| `linux` | `.venv/bin/python3` | `apt` / `dnf` 等 | ディストリ次第（通常 PATH 上に入る） |
| `windows` | `.venv/Scripts/python.exe` | winget／公式インストーラ | `/c/Program Files/LibreOffice/program/soffice.exe` |

> **`PYBIN` の求め方**（Bash 間で変数は持ち越されないため、各ステップの先頭で毎回求め直す）：
> ```bash
> PYBIN="$ROOT/.venv/bin/python3"; [ -x "$PYBIN" ] || PYBIN="$ROOT/.venv/Scripts/python.exe"
> ```
> mac/Linux は `bin/python3`、Windows は `Scripts/python.exe` を指す（MSYS は `.exe` を実行可能扱いするため `-x` 判定でよい）。

### 2. システム依存のチェック（pip 外・ホスト側）

次を実行して現状を把握する（mac / Linux / Windows 共通）：

```bash
HOSTPY="$(command -v python3 || command -v python || true)"
if [ -n "$HOSTPY" ]; then echo "python:  $HOSTPY ($("$HOSTPY" --version 2>&1))"; else echo "python:  NOT FOUND"; fi
echo "git:     $(git --version 2>&1)"
printf 'soffice: '
soffice_path="$(command -v soffice || true)"
if [ -z "$soffice_path" ]; then
  for c in \
    "/Applications/LibreOffice.app/Contents/MacOS/soffice" \
    "/c/Program Files/LibreOffice/program/soffice.exe" \
    "/c/Program Files (x86)/LibreOffice/program/soffice.exe"; do
    [ -x "$c" ] && soffice_path="$c" && break
  done
fi
echo "${soffice_path:-NOT FOUND}"
printf 'pkg-mgr: '; (command -v brew || command -v winget || echo 'brew/winget どちらも無し')
```

判定基準と対応（`OSKIND` に応じて導入コマンドを選ぶ）：

| 項目 | 必須度 | 不足時の対応 |
|---|---|---|
| Python ≥ 3.10 | 必須 | mac: `brew install python@3.12`／Windows: `winget install --id Python.Python.3.12`（または [python.org 公式インストーラ](https://www.python.org/downloads/windows/)。"Add python.exe to PATH" にチェック）／Linux: ディストリのパッケージ。3.10 未満では venv を作らない（`slides.py` が PEP 604 記法・`Pillow>=12.3` を使用するため import で落ちる） |
| `git` | 必須 | 通常は導入済み。無ければ mac: `xcode-select --install`／Windows: `winget install --id Git.Git`（または [Git for Windows](https://git-scm.com/downloads/win)）／Linux: `apt install git` 等 |
| `soffice` | 必須 | §2-a 参照 |
| `cairo`（＋`libffi`） | 任意（`add_icon` 使用時のみ） | アイコンを使わないなら未導入のままでよい。入れるなら mac: `brew install cairo libffi`／Linux: `apt install libcairo2 libffi-dev python3-dev`（Debian 系）等／Windows: [GTK for Windows ランタイム](https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer)で cairo を導入（[CairoSVG 公式](https://cairosvg.org/documentation/)準拠。手間なので当面はスキップ推奨）。詳細は README §0-2「cairo の入れ方」参照 |
| パッケージマネージャ | 導入の前提 | mac: [brew.sh](https://brew.sh)／Windows: winget（Win10/11 標準同梱。無ければ各公式インストーラで代替）／Linux: ディストリ標準 |

**システムを変更する導入コマンド（`brew install` / `winget install` 系）は、実行前に「これを実行してよいか」をユーザーに確認してから走らせる。** 大きなダウンロード（LibreOffice）を含むため勝手に始めない。ユーザーが「自分でやる」と言えばコマンドだけ提示して待つ。

> **winget を非対話で走らせる場合**は `--accept-source-agreements --accept-package-agreements` を付ける（初回のソース同意プロンプトで止まるのを防ぐ）。例：`winget install --id TheDocumentFoundation.LibreOffice --accept-source-agreements --accept-package-agreements`。

#### 2-a. `soffice` を PATH に通す

`slides.py` は `shutil.which("soffice")` で soffice を探す（fallback は mac の Homebrew パスのみ）。したがって**どの OS でも soffice が PATH 上にあること**が必要。`command -v soffice` が出れば OK。出ない場合は OS 別に対応する：

- **macOS**：cask 版は `soffice` を PATH に登録せず、実体を `/Applications/LibreOffice.app/Contents/MacOS/soffice` に置く。次のどちらかで通す（A は要確認のうえ実行可、B はユーザー作業）：
  - A. シンボリックリンク：`ln -s /Applications/LibreOffice.app/Contents/MacOS/soffice "$(brew --prefix)/bin/soffice"`
  - B. PATH 追加：`~/.zshrc` 等に `export PATH="/Applications/LibreOffice.app/Contents/MacOS:$PATH"`
  - どちらの実体も無ければ `brew install --cask libreoffice`（要確認）。
- **Windows**：実体は `C:\Program Files\LibreOffice\program\`（`soffice.exe` / `soffice.com`）。既定では PATH に入らないので、**ユーザー作業として**「設定 → システム → バージョン情報 → システムの詳細設定 → 環境変数」でユーザーの `Path` に `C:\Program Files\LibreOffice\program` を追加し、**Claude Code を開き直す**（Bash ツールは起動時の Windows PATH を継承するため、再起動しないと反映されない）。未導入なら `winget install --id TheDocumentFoundation.LibreOffice`（要確認）→ その後 PATH 追加。Windows ではシンボリックリンクでの回避は勧めない。
- **Linux**：パッケージ導入で通常 PATH 上に入る。入らなければ実体（例：`/usr/bin/soffice`）の場所を確認して PATH を通す。

設定後、`command -v soffice` がパスを返すことを必ず再確認する。

`soffice` は手順 5 の検査ゲート（`validate_render`）と視覚確認（`render_pngs`）の必須依存。ここで揃わないと `/slides-gen` が手順 5 で停止するため、**未解決のまま「セットアップ完了」と報告しない**。

### 3. venv 作成

`.venv/` が無ければ作る（あれば再利用）。Python 依存はプロジェクト内に閉じるので確認不要で進めてよい。

```bash
ROOT="$(git rev-parse --show-toplevel)"
PYBIN="$ROOT/.venv/bin/python3"; [ -x "$PYBIN" ] || PYBIN="$ROOT/.venv/Scripts/python.exe"
if [ -x "$PYBIN" ]; then
  echo "venv exists"
else
  HOSTPY="$(command -v python3 || command -v python)"
  "$HOSTPY" -m venv "$ROOT/.venv"
  # 作成後に PYBIN を求め直す（OS により bin/ か Scripts/ か決まる）
  PYBIN="$ROOT/.venv/bin/python3"; [ -x "$PYBIN" ] || PYBIN="$ROOT/.venv/Scripts/python.exe"
fi
"$PYBIN" --version
```

### 4. 依存インストール

本番依存を入れる（`pip` は venv 内の Python 経由で明示呼び出し）：

```bash
ROOT="$(git rev-parse --show-toplevel)"
PYBIN="$ROOT/.venv/bin/python3"; [ -x "$PYBIN" ] || PYBIN="$ROOT/.venv/Scripts/python.exe"
"$PYBIN" -m pip install --quiet --upgrade pip
"$PYBIN" -m pip install -r "$ROOT/requirements.txt"
```

引数に `--with-tests` が指定された場合、またはユーザーが回帰テストを回したい場合は、本番依存とあわせて開発依存も入れる：

```bash
ROOT="$(git rev-parse --show-toplevel)"
PYBIN="$ROOT/.venv/bin/python3"; [ -x "$PYBIN" ] || PYBIN="$ROOT/.venv/Scripts/python.exe"
"$PYBIN" -m pip install -r "$ROOT/requirements.txt" -r "$ROOT/requirements-dev.txt"
```

> `requirements.txt` の `cairosvg`/`requests` はアイコン使用時のみ実際に必要。pip では入るが、ランタイムで `add_icon` を使う場合は §2 の `cairo`（OS 側）も必要になる点に留意（Windows は別途 `libcairo-2.dll` の用意が必要）。

### 5. 検証

`/slides-gen` が前提とする 4 点（Python・soffice・PyMuPDF・ライブラリ本体）が揃ったか確認する：

```bash
ROOT="$(git rev-parse --show-toplevel)"
PYBIN="$ROOT/.venv/bin/python3"; [ -x "$PYBIN" ] || PYBIN="$ROOT/.venv/Scripts/python.exe"
"$PYBIN" --version \
  && (command -v soffice \
       || ls "/Applications/LibreOffice.app/Contents/MacOS/soffice" \
       || ls "/c/Program Files/LibreOffice/program/soffice.exe") \
  && "$PYBIN" -c "import fitz; print('PyMuPDF OK')" \
  && (cd "$ROOT" && "$PYBIN" -c "import slides; print('slides OK')")
```

> `command -v soffice` が空で、既定パスにのみ存在する場合は「導入済みだが PATH 未登録」。§2-a に戻って PATH を通す（Windows は Claude Code の再起動が必要）。PATH に無いと `/slides-gen` の `shutil.which("soffice")` が失敗して手順 5 で止まる。

`--with-tests` 時は続けて回帰テストも通す：

```bash
ROOT="$(git rev-parse --show-toplevel)"
PYBIN="$ROOT/.venv/bin/python3"; [ -x "$PYBIN" ] || PYBIN="$ROOT/.venv/Scripts/python.exe"
(cd "$ROOT" && "$PYBIN" -m pytest -q)
```

いずれかが失敗したら原因（依存欠落・PATH 等）を切り分け、§2〜4 の該当ステップに戻る。失敗を握り潰して「完了」と報告しない。

### 6. 完了報告

次の形で報告する：

- **揃ったもの**：Python 版／venv の Python パス（`PYBIN`）／soffice パス／PyMuPDF／`slides`（＋テストを回した場合は結果）。
- **ユーザー対応が必要なもの**（あれば）：未導入のシステム依存と、その OS 向けの導入コマンド（mac: `brew install ...`／Windows: `winget install ...`／Linux: 各パッケージ）。Windows で soffice を入れた場合は PATH 追加＋ Claude Code の再起動が要る点も。アイコンを使う予定なら `cairo`（Windows は `libcairo-2.dll`）も。
- **次のステップ**：`/slides-new` で資料を作成 → `/slides-gen` で生成、と案内する。

> このスキルは環境構築までを担当する。スライド生成そのものは `/slides-new` → `/slides-gen` の役割。
