"""<config> タグの解析・句読点の表示名・多言語化の文言のテスト

<config> の解析は今は gui.App にある (フェーズ0b で settings.py に移す予定)。
移したときにこのテストの入出力の例をそのまま使えるようにしておく。
"""

import re

import pytest

gui = pytest.importorskip("gui")
import i18n  # noqa: E402

CONFIG_EXAMPLES = [
    # (ノートのリスト, 解析結果)
    (["<config speaker=四国めたん style=ノーマル pause=0.5>"],
     {"speaker": "四国めたん", "style": "ノーマル", "pause": "0.5"}),
    (['<config font="Arial Rounded MT Bold" fontsize=26>'],
     {"font": "Arial Rounded MT Bold", "fontsize": "26"}),
    (['<config font="">'], {"font": ""}),
    (["本文\n<config engine=openai model=kokoro speaker=af_heart split_sentences=on>\n続き"],
     {"engine": "openai", "model": "kokoro", "speaker": "af_heart", "split_sentences": "on"}),
    # 複数のスライドにある場合は後のタグが優先
    (["<config speed=1.0 pitch=0.00>", "<config speed=1.2>"], {"speed": "1.2", "pitch": "0.00"}),
    (["<CONFIG kuten=そのまま touten=、>"], {"kuten": "そのまま", "touten": "、"}),
    (["タグなし"], {}),
]


@pytest.mark.parametrize("notes,expected", CONFIG_EXAMPLES)
def test_parse_config_tags(notes, expected):
    assert gui.App._parse_config_tags(gui.App, notes) == expected


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
