"""共有 fixture。リポジトリルートの import パスは pyproject の
`[tool.pytest.ini_options] pythonpath` で解決する。"""
import pytest

from slides import load_template, L_BODY


@pytest.fixture
def prs():
    """雛形を読み込んだ Presentation（検知レジストリもリセット済み）。"""
    return load_template()


@pytest.fixture
def body_slide(prs):
    """本文レイアウト（Layout 2）のスライドを 1 枚追加して返す。"""
    return prs.slides.add_slide(prs.slide_layouts[L_BODY])
