"""notes.py (ノート記法の解析) のテスト

今の動作を固定するためのテスト。動作を意図して変えたときは、テストも合わせて直す。
"""

from notes import (
    _BRACE_PATTERN,
    _NEXT_TAG,
    _READING_PATTERN,
    _split_sentences,
    _to_display,
    _to_reading,
    parse_notes,
)

# 字幕テキスト中のプレースホルダ
LT, GT, MATH_START, MATH_END, LBRACE, RBRACE = "\x02", "\x03", "\x13", "\x14", "\x15", "\x16"


def sentences(text, split=False):
    return [(s.display, s.reading) for s in parse_notes(text, split).sentences]


# ---------------------------------------------------------------------------
# 文の区切り
# ---------------------------------------------------------------------------

def test_split_by_lines():
    assert sentences("一行目。\n二行目。") == [("一行目。", "一行目。"), ("二行目。", "二行目。")]


def test_blank_lines_are_ignored():
    assert sentences("一行目。\n\n  \n二行目。") == [("一行目。", "一行目。"), ("二行目。", "二行目。")]


def test_br_keeps_sentence_and_breaks_subtitle():
    assert sentences("前半<br>後半") == [("前半\n後半", "前半後半")]


def test_wait_splits_and_sets_pause():
    p = parse_notes("前。<wait=1s>後。\n次。<wait=500ms>\n最後。")
    assert [s.display for s in p.sentences] == ["前。", "後。", "次。", "最後。"]
    assert p.pauses == [1.0, None, 0.5]


def test_leading_wait():
    p = parse_notes("<wait=2>最初。")
    assert p.leading_pause == 2.0
    assert [s.display for s in p.sentences] == ["最初。"]


def test_config_tag_is_removed():
    p = parse_notes("<config speaker=x fontsize=20>\n本文。")
    assert [s.display for s in p.sentences] == ["本文。"]


def test_tag_only_line_is_not_a_sentence():
    # v0.7.3: 読み上げる文字がない行は文にしない (以前は空の文になり、Kokoro ではエラーになっていた)
    # その行の音声タグは次の行には効かない
    p = parse_notes("一行目。\n<speed=1.2>\n二行目。")
    assert [(s.display, s.reading, s.speed) for s in p.sentences] == [
        ("一行目。", "一行目。", None), ("二行目。", "二行目。", None)]
    assert p.pauses == [None]


def test_format_tag_only_line_is_not_a_sentence():
    assert [s.display for s in parse_notes("<b></b>\n本文。").sentences] == ["本文。"]


def test_tag_only_line_keeps_wait_and_next():
    p = parse_notes("前。<wait=1s>\n<volume=0.5>\n後。")
    assert [s.display for s in p.sentences] == ["前。", "後。"]
    assert p.pauses == [1.0]
    p = parse_notes("一行目。\n<next><speed=1.2>\n二行目。")
    assert p.next_positions == [(0, 1.0)]


def test_tag_on_same_line_still_applies():
    assert parse_notes("一行目。\n<speed=1.2>二行目。").sentences[1].speed == 1.2


def test_only_tags_gives_no_sentences():
    assert parse_notes("<speed=1.2>").sentences == []


# ---------------------------------------------------------------------------
# 文末 (. ! ?) での区切り
# ---------------------------------------------------------------------------

def test_split_sentence_ends_off_by_default():
    assert len(parse_notes("First. Second.").sentences) == 1


def test_split_sentence_ends():
    text = "Welcome to {PPVoice|P P Voice}. This is it! Isn't it great? Yes."
    assert [d for d, _ in sentences(text, split=True)] == [
        "Welcome to PPVoice.", "This is it!", "Isn't it great?", "Yes."]


def test_split_sentence_ends_keeps_abbreviations():
    text = "See Fig. 3 and Eq. 2. Dr. Smith and J. Doe agree, e.g. it works. 1. Intro. Pi is 3.14 here."
    assert [d for d, _ in sentences(text, split=True)] == [
        "See Fig. 3 and Eq. 2.", "Dr. Smith and J. Doe agree, e.g. it works.", "1. Intro.", "Pi is 3.14 here."]


def test_split_sentence_ends_does_not_split_inside_braces():
    text = "{e.g.|for example} this works. Next."
    # {...} 内の句読点は句読点置換の対象外にするため、仮の記号 ( = ".") になる
    assert [d for d, _ in sentences(text, split=True)] == ["eg this works.", "Next."]


def test_split_sentence_ends_with_next():
    p = parse_notes("First point.<next> Second point. <next>Third point.", split_sentence_ends=True)
    assert [s.display for s in p.sentences] == ["First point.", "Second point.", "Third point."]
    assert p.next_positions == [(0, 1.0), (2, 0.0)]


def test_japanese_text_is_not_split_at_ascii_period():
    assert len(parse_notes("英語の文 This is a pen. はここ。", split_sentence_ends=True).sentences) == 1


# ---------------------------------------------------------------------------
# 読み指定
# ---------------------------------------------------------------------------

def test_reading():
    assert sentences("{PPTX|パワーポイント}ファイル") == [("PPTXファイル", "パワーポイントファイル")]


def test_accent():
    s = parse_notes("{橋|はし|2}を渡る。{箸|はし|1}を持つ。").sentences[0]
    assert s.accents == [("ハシ", 2), ("ハシ", 1)]
    assert s.reading == "はしを渡る。はしを持つ。"


def test_brace_without_reading_is_protected():
    s = parse_notes("{<b>}はタグではない").sentences[0]
    assert s.display == f"{LT}b{GT}はタグではない"


def test_single_braces_are_kept():
    assert sentences("a { b") == [("a { b", "a { b")]
    assert sentences("a } b") == [("a } b", "a } b")]


def test_braces_without_bar_show_content():
    # {...} (| なし) は中身をそのまま表示する記法なので、波括弧は消える
    assert sentences("a { b } c") == [("a  b  c", "a  b  c")]


def test_adjacent_readings_with_digits():
    # 不具合の再現ケース: 仮の記号の境界を誤って置換していた (v0.6.1 で修正)
    text = "{数値|すうち}と{固有値|こゆうち}\n2つ目{は|わ}、2{行|ぎょう}2{列|れつ}の行列"
    assert sentences(text)[1] == ("2つ目は、2行2列の行列", "2つ目わ、2ぎょう2れつの行列")


def test_many_placeholders():
    text = "".join(f"{i}{{a{i}|b{i}}}" for i in range(15))
    raw, _, _, _, _ = _split_sentences(text)
    assert raw == [text]


def test_punctuation_in_braces_is_escaped():
    # {...} 内の句読点は句読点置換の対象外 (プレースホルダに置き換わる)
    s = parse_notes("{A。B|えーびー}、終わり。").sentences[0]
    assert "。" not in s.display.split("、")[0]


# ---------------------------------------------------------------------------
# 数式
# ---------------------------------------------------------------------------

def test_math():
    s = parse_notes(r"式 {$\frac{a}{b} < |c|$|ぶんすう} です").sentences[0]
    assert s.display == f"式 {MATH_START}\\frac{LBRACE}a{RBRACE}{LBRACE}b{RBRACE} {LT} |c|{MATH_END} です"
    assert s.reading == "式 ぶんすう です"


def test_dollar_outside_reading_is_plain_text():
    assert sentences("価格は$100です") == [("価格は$100です", "価格は$100です")]


def test_mixed_dollar_text_is_not_math():
    s = parse_notes("{$a$ と $b$|えーとびー}").sentences[0]
    assert MATH_START not in s.display


# ---------------------------------------------------------------------------
# タグ
# ---------------------------------------------------------------------------

def test_format_tags_stay_in_display_and_are_removed_from_reading():
    s = parse_notes("<b>太字</b>と<color=#FF0000>赤</color>").sentences[0]
    assert s.display == "<b>太字</b>と<color=#FF0000>赤</color>"
    assert s.reading == "太字と赤"


def test_voice_tags():
    s = parse_notes("<speed=1.3><pitch=0.05><intonation=0.5><volume=0.8>文").sentences[0]
    assert (s.speed, s.pitch, s.intonation, s.volume) == (1.3, 0.05, 0.5, 0.8)
    assert s.display == "文"
    assert s.reading == "文"


def test_last_voice_tag_wins():
    assert parse_notes("<speed=1.1>あ<speed=1.5>い").sentences[0].speed == 1.5


def test_unknown_tag_is_kept():
    # 現状の動作: 知らないタグは字幕にも読みにも残る (書き間違いの検出で警告する予定)
    s = parse_notes("<sped=1.2>速く読む。").sentences[0]
    assert s.display == "<sped=1.2>速く読む。"
    assert s.reading == "<sped=1.2>速く読む。"


# ---------------------------------------------------------------------------
# <next>
# ---------------------------------------------------------------------------

def test_next_between_lines():
    p = parse_notes("最初。\n<next>\n次。")
    assert [s.display for s in p.sentences] == ["最初。", "次。"]
    assert p.next_positions == [(0, 1.0)]


def test_next_at_start():
    assert parse_notes("<next>最初。").next_positions == [(0, 0.0)]


def test_next_alone_at_start():
    assert parse_notes("<next>\n最初。").next_positions == [(-1, 0.0)]


def test_next_mid_sentence():
    p = parse_notes("ここで<next>出ます")
    assert p.next_positions == [(0, 3 / 6)]  # 「ここで」3文字 / 「ここで出ます」6文字


def test_next_inside_braces_is_not_counted():
    p = parse_notes("{<next>|ネクストタグ}を書く")
    assert p.next_positions == []
    assert p.sentences[0].reading == "ネクストタグを書く"


def test_patterns_used_by_next_check():
    # GUI の「<next>確認」は {…} を除いてから <next> を数える
    text = "{<next>|ネクスト}と<next>と{<next>}"
    stripped = _BRACE_PATTERN.sub("", _READING_PATTERN.sub("", text))
    assert len(_NEXT_TAG.findall(stripped)) == 1


# ---------------------------------------------------------------------------
# 低レベル関数
# ---------------------------------------------------------------------------

def test_to_display_and_reading():
    assert _to_display("{漢字|かんじ}<wait=1s>") == "漢字"
    assert _to_reading("{漢字|かんじ}<b>x</b>") == "かんじx"
