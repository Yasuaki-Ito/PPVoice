"""読みの辞書 (<dict>) のテスト"""

from notes import apply_dictionary, parse_notes
from settings import Settings, parse_dict_tags

D = {"PPTX": "パワーポイント", "GUI": "ジーユーアイ", "端": "はし|0", "Fig.": "フィギュア",
     "Power": "パワー", "PowerPoint": "パワーポイント"}


def test_dictionary_replaces_words():
    text, hits = apply_dictionary("PPTX と 端", D)
    assert text == "{PPTX|パワーポイント} と {端|はし|0}"
    assert hits == {"PPTX": 1, "端": 1}


def test_dictionary_prefers_longer_words():
    assert apply_dictionary("PowerPoint と Power", D)[0] == "{PowerPoint|パワーポイント} と {Power|パワー}"


def test_dictionary_ascii_word_boundaries():
    assert apply_dictionary("GUIDE と GUI_x と GUIの", D)[0] == "GUIDE と GUI_x と {GUI|ジーユーアイ}の"
    assert apply_dictionary("gui", D)[0] == "gui"  # 大文字・小文字は区別する


def test_dictionary_skips_readings_braces_tags_and_math():
    text = "{端|はた} {GUI} <font=PowerPoint>x</font> {$GUI$|じーゆーあい} <speaker=GUI>"
    assert apply_dictionary(text, D) == (text, {})


def test_dictionary_in_parse_notes():
    p = parse_notes("端を渡る。Fig. 2 を見る。", True, D)
    assert [(s.reading, s.accents) for s in p.sentences] == [("はしを渡る。フィギュア 2 を見る。", [("ハシ", 0)])]
    assert p.dict_hits == {"端": 1, "Fig.": 1}


def test_dict_tag_is_not_read():
    assert [s.display for s in parse_notes("<dict PPTX=パワーポイント>\n本文。").sentences] == ["本文。"]


def test_parse_dict_tags_and_roundtrip():
    d = parse_dict_tags(['<dict PPTX=パワーポイント 端=はし|0 "Fig."=フィギュア>', '<DICT "a b"="エー ビー" PPTX=ピーピーティーエックス>'])
    assert d == {"PPTX": "ピーピーティーエックス", "端": "はし|0", "Fig.": "フィギュア", "a b": "エー ビー"}
    s = Settings(dictionary=d)
    assert parse_dict_tags([s.to_dict_tag()]) == d
    assert Settings().to_dict_tag() == ""
