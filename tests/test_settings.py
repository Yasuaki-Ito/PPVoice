"""settings.py (設定と <config> タグ) と pipeline の話者の解決のテスト"""

import pytest

from pipeline import resolve_openai_voice, resolve_voicevox_style
from settings import Settings, parse_config_tags

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
    assert parse_config_tags(notes) == expected


def test_default_config_tag():
    # 以前の GUI の「設定保存」と同じ並び・桁数 (話者を選んでいない初期状態)
    assert Settings().to_config_tag() == (
        "<config engine=voicevox pause=0.5 speed=1.0 pitch=0.00 intonation=1.0 volume=1.0 end_pause=2.0 "
        "split_sentences=off auto_next=5.0 auto_next_enabled=on subtitle=on subtitle_style=outline fontsize=18 "
        'font="" bottom=0.05 font_color=#FFFFFF outline=on outline_color=#000000 outline_width=0.75 glow=off '
        "glow_color=#000000 glow_size=11.0 bg_color=#000000 bg_alpha=60 kuten=そのまま touten=そのまま bold=off "
        "italic=off underline=off math_bold=on>")


def test_config_roundtrip():
    s = Settings()
    s.apply_config(parse_config_tags([
        '<config speaker=四国めたん style=ノーマル pause=0.4 speed=1.2 fontsize=26 font="HGP創英角ﾎﾟｯﾌﾟ体" '
        "font_color=#ffff00 glow=on glow_size=8.0 kuten=．(全角) bold=on math_bold=off>"]))
    s2 = Settings()
    s2.apply_config(parse_config_tags([s.to_config_tag()]))
    assert s2 == s
    assert s.font_color == "#FFFF00"  # 色は大文字にそろえる
    assert s.kuten == "．(全角)"


def test_openai_config_tag_has_model_and_no_style():
    s = Settings()
    s.apply_config({"engine": "openai", "speaker": "af_heart", "style": "ノーマル"})
    tag = s.to_config_tag()
    assert tag.startswith("<config engine=openai model=kokoro speaker=af_heart pause=")
    assert " style=" not in tag
    assert "api_key" not in tag


def test_engine_switch_resets_speaker_and_split():
    s = Settings()
    s.speaker, s.style = "ずんだもん", "ノーマル"
    s.set_engine("openai")
    assert (s.speaker, s.style, s.split_sentences) == ("", "", True)
    s.set_engine("voicevox")
    assert s.split_sentences is False


def test_config_split_sentences_overrides_engine_default():
    s = Settings()
    s.apply_config({"engine": "openai", "split_sentences": "off"})
    assert s.split_sentences is False


def test_speaker_without_style_resets_style():
    s = Settings()
    s.speaker, s.style = "四国めたん", "あまあま"
    s.apply_config({"speaker": "ずんだもん"})
    assert (s.speaker, s.style) == ("ずんだもん", "")


def test_invalid_value_is_skipped():
    warned = []
    s = Settings()
    s.apply_config({"speed": "abc", "fontsize": "26"}, warn=lambda k, v: warned.append(k))
    assert (s.speed, s.fontsize, warned) == (1.0, 26, ["speed"])


def test_urls_per_engine():
    s = Settings()
    s.url = "http://example:50021"
    s.set_engine("openai")
    assert s.url == "http://localhost:8880/v1"
    s.set_engine("voicevox")
    assert s.url == "http://example:50021"


def test_embed_kwargs():
    s = Settings()
    s.font = "メイリオ"
    s.auto_next_enabled = False
    kw = s.embed_kwargs()
    assert kw["subtitle_font_name"] == "メイリオ"
    assert kw["subtitle_font_color"] == "FFFFFF"
    assert kw["auto_next_interval_ms"] == -1
    assert kw["end_pause_ms"] == 2000


# ---------------------------------------------------------------------------
# 話者名 → ID の解決
# ---------------------------------------------------------------------------

SPEAKERS = [
    {"name": "四国めたん", "styles": [{"name": "ノーマル", "id": 2}, {"name": "あまあま", "id": 0}]},
    {"name": "ずんだもん", "styles": [{"name": "ノーマル", "id": 3}, {"name": "ささやき", "id": 22}]},
    {"name": "スタイルなし", "styles": []},
]


def test_resolve_voicevox_style():
    assert resolve_voicevox_style(SPEAKERS, "ずんだもん", "ささやき") == (22, "ずんだもん", "ささやき")
    assert resolve_voicevox_style(SPEAKERS, "ずんだもん", "") == (3, "ずんだもん", "ノーマル")


def test_resolve_voicevox_style_fallbacks(capsys):
    # 見つからない話者 → 一覧の最初の話者、見つからないスタイル → その話者の最初のスタイル
    assert resolve_voicevox_style(SPEAKERS, "いない人", "ノーマル") == (2, "四国めたん", "ノーマル")
    assert resolve_voicevox_style(SPEAKERS, "ずんだもん", "ない") == (3, "ずんだもん", "ノーマル")
    assert resolve_voicevox_style(SPEAKERS, "", "") == (2, "四国めたん", "ノーマル")
    out = capsys.readouterr().out
    assert "いない人" in out and "ない" in out


def test_resolve_openai_voice():
    assert resolve_openai_voice(["af_heart", "am_adam"], "am_adam") == "am_adam"
    assert resolve_openai_voice(["af_heart", "am_adam"], "") == "af_heart"
