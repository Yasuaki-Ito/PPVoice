"""pptx_writer の字幕の書式のテスト"""

from pptx_writer import _parse_formatted_text


def segments(text):
    return [(s.text, s.color, s.size, s.font, s.bold) for s in _parse_formatted_text(text)]


def test_format_tags():
    assert segments("<b>太</b><color=#ff0000>赤</color><size=+2>大</size><font=メイリオ>字</font>") == [
        ("太", None, None, None, True), ("赤", "FF0000", None, None, False),
        ("大", None, "+2", None, False), ("字", None, None, "メイリオ", False)]


def test_invalid_values_are_shown_as_text():
    # 以前は <color=red> で生成が止まっていた (int('', 16))
    assert segments("<color=red>赤</color>") == [("<color=red>赤", None, None, None, False)]
    assert segments("<size=x>a<b=1>b") == [("<size=x>a<b=1>b", None, None, None, False)]
