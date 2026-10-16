"""句読点の表示名・多言語化の文言のテスト

<config> の解析のテストは test_settings.py にある。
"""

import re

import pytest

gui = pytest.importorskip("gui")
import i18n  # noqa: E402

# ---------------------------------------------------------------------------
# 句読点の置換先: 内部の値 (<config> に保存する値) は日本語のまま、表示だけ翻訳する
# ---------------------------------------------------------------------------

@pytest.fixture
def english():
    i18n.set_language("en")
    yield
    i18n.set_language("ja")


def test_punct_label_and_value_roundtrip(english):
    for choices in (gui._KUTEN_CHOICES, gui._TOUTEN_CHOICES):
        for value, _ in choices:
            label = gui.App._punct_label(choices, value)
            assert gui.App._punct_value(choices, label) == value


def test_punct_label_in_english(english):
    assert gui.App._punct_label(gui._KUTEN_CHOICES, "そのまま") == "Unchanged"
    assert gui.App._punct_label(gui._KUTEN_CHOICES, "。") == "。"
    # 英語の表示名で書かれた値も受け付ける
    assert gui.App._punct_value(gui._KUTEN_CHOICES, "Unchanged") == "そのまま"


# ---------------------------------------------------------------------------
# 多言語化の文言
# ---------------------------------------------------------------------------

def test_i18n_placeholders_match():
    fmt = re.compile(r"{(\w+)}")
    for key, (ja, en) in i18n.STRINGS.items():
        assert set(fmt.findall(ja)) == set(fmt.findall(en)), key


def test_i18n_keys_used_in_code_exist():
    import os
    src = os.path.join(os.path.dirname(__file__), "..", "src")
    code = ""
    for root, _, files in os.walk(src):
        for f in files:
            if f.endswith(".py"):
                code += open(os.path.join(root, f), encoding="utf-8").read()
    used = set(re.findall(r'\bt\("([a-z0-9_]+)"', code))
    assert used - set(i18n.STRINGS) == set()
