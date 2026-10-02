"""ノート欄の記法の解析

読み指定 ({表示|読み}), 書式タグ, 音声制御タグ (<speed> など), <wait>, <next>, <br>, 数式
を解析し、字幕用テキスト・TTS に渡す読み・文ごとのパラメータに分解する。
TTS エンジンに依存しない処理なので、エンジン (tts/), 字幕 (pptx_writer.py), GUI から共通に使う。
"""

import re
from dataclasses import dataclass, field

# 読み指定パターン: {表示テキスト|読み} or {表示テキスト|読み|アクセント位置}
# 表示テキストが $LaTeX$ の場合は数式として扱う ({} や | を含んでもよい)
_READING_PATTERN = re.compile(r"\{(\$(?:[^$\\]|\\.)+\$|[^|}]+)\|([^|}]+)(?:\|(\d+))?\}")
# 数式の表示テキスト: $LaTeX$
_MATH_DISPLAY = re.compile(r"\$((?:[^$\\]|\\.)+)\$")
# 保護パターン: {テキスト} (|なし) — 文分割を抑制
_BRACE_PATTERN = re.compile(r"\{([^|}]+)\}")

# 書式タグパターン: <b>, </b>, <i>, </i>, <u>, </u>, <color=#RRGGBB>, </color>,
# <font=...>, </font>, <size=N>, </size>, <br>, <wait=Ns>,
# <speed=N>, <pitch=N>, <intonation=N>, <volume=N>, <config ...>
_FORMAT_TAG = re.compile(
    r"</?(?:b|i|u|color(?:=#[0-9a-fA-F]{6})?|font(?:=[^>]+)?|size(?:=[+\-]?\d+)?"
    r"|speed(?:=[\d.]+)?|pitch(?:=[+\-]?[\d.]+)?"
    r"|intonation(?:=[\d.]+)?|volume(?:=[\d.]+)?)>"
    r"|<br\s*/?>|<wait=[\d.]+\s*(?:ms|s)?\s*>|<config\s[^>]*>|<next\s*/?>",
    re.IGNORECASE,
)

# <config ...> タグ
_CONFIG_TAG = re.compile(r"<config\s[^>]*>", re.IGNORECASE)

# <br> 分割パターン
_SPLIT_BR = re.compile(r"<br\s*/?>", re.IGNORECASE)

# <wait=Ns> パターン (数値+単位キャプチャ)
# 対応形式: <wait=1s>, <wait=0.5s>, <wait=500ms>, <wait=2> (単位なし=秒)
_WAIT_TAG = re.compile(r"<wait=([\d.]+)\s*(ms|s)?\s*>", re.IGNORECASE)

# <speed=N> / <pitch=N> / <intonation=N> / <volume=N> パターン (音声パラメータ)
_SPEED_TAG = re.compile(r"<speed=([\d.]+)>", re.IGNORECASE)
_PITCH_TAG = re.compile(r"<pitch=([+\-]?[\d.]+)>", re.IGNORECASE)
_INTONATION_TAG = re.compile(r"<intonation=([\d.]+)>", re.IGNORECASE)
_VOLUME_TAG = re.compile(r"<volume=([\d.]+)>", re.IGNORECASE)

# <next> パターン (アニメーション発火)
_NEXT_TAG = re.compile(r"<next\s*/?>", re.IGNORECASE)

# _split_sentences で {...} を保護するプレースホルダ: \x00番号\x00
_PLACEHOLDER = re.compile("\x00(\\d+)\x00")

# プレースホルダ: {テキスト} 内の文字をエスケープするための代替文字
_LT = "\x02"
_GT = "\x03"
# 数式マーカー: 字幕テキスト中の LaTeX を囲む (pptx_writer.py で数式に変換)
_MATH_START = "\x13"
_MATH_END = "\x14"
# 数式内の {} ({テキスト} パターンとして展開されないよう退避)
_LBRACE = "\x15"
_RBRACE = "\x16"
# 句読点プレースホルダ ({...} 内の句読点を置換対象外にする)
_PUNCT_PH = {
    "。": "\x04", "．": "\x05", ".": "\x06",
    "、": "\x07", "，": "\x10", ",": "\x11",
}


def _hira_to_kata(text: str) -> str:
    """ひらがなをカタカナに変換する。"""
    return "".join(
        chr(ord(c) + 0x60) if "\u3041" <= c <= "\u3096" else c
        for c in text
    )


def _extract_accents(text: str) -> list[tuple[str, int]]:
    """文中の {表示|読み|N} からアクセント指定を抽出する。

    Returns: [(カタカナ読み, accent_position), ...]
    """
    accents = []
    for m in _READING_PATTERN.finditer(text):
        if m.group(3) is not None:
            katakana = _hira_to_kata(m.group(2))
            accents.append((katakana, int(m.group(3))))
    return accents


def _to_display(text: str) -> str:
    """{表示|読み} → 表示, {テキスト} → テキスト に変換 (字幕用)。

    <br> は改行文字に変換する。
    {テキスト} 内の <> はプレースホルダに変換し、
    書式タグとして解釈されないようにする。
    """
    # {表示|読み} / {テキスト} 内の <> をエスケープして展開
    # (エスケープしないと中の制御タグが除去されてしまう)
    def _escape_content(m):
        s = m.group(1).replace("<", _LT).replace(">", _GT)
        for ch, ph in _PUNCT_PH.items():
            s = s.replace(ch, ph)
        return s

    def _escape_reading(m):
        mm = _MATH_DISPLAY.fullmatch(m.group(1))
        if mm:
            latex = (mm.group(1).replace("<", _LT).replace(">", _GT)
                     .replace("{", _LBRACE).replace("}", _RBRACE))
            return f"{_MATH_START}{latex}{_MATH_END}"
        return _escape_content(m)
    text = _READING_PATTERN.sub(_escape_reading, text)
    text = _BRACE_PATTERN.sub(_escape_content, text)
    # <wait>, <config>, <speed>, <pitch> タグを除去 (エスケープ済みのものはマッチしない)
    text = _WAIT_TAG.sub("", text)
    text = _CONFIG_TAG.sub("", text)
    text = _SPEED_TAG.sub("", text)
    text = _PITCH_TAG.sub("", text)
    text = _INTONATION_TAG.sub("", text)
    text = _VOLUME_TAG.sub("", text)
    text = _NEXT_TAG.sub("", text)
    # <br> → 改行 (エスケープ済みのものはマッチしない)
    return _SPLIT_BR.sub("\n", text)


def _to_reading(text: str) -> str:
    """{表示|読み} → 読み, {テキスト} → テキスト に変換 (TTS用)。

    書式タグは読み上げに不要なため除去する。
    """
    text = _READING_PATTERN.sub(r"\2", text)
    text = _BRACE_PATTERN.sub(r"\1", text)
    return _FORMAT_TAG.sub("", text)


# 文末 (. ! ?) で区切る位置: 閉じ引用符・括弧 (と直後の <next>) の後に空白があり、
# 次が大文字・数字・開き引用符・タグ・プレースホルダ ({...} / <next>) で始まる
_SENTENCE_END = re.compile("(?<=[.!?])([\"'”’)\\]]*\x12*)[ \t]+(?=[A-Z0-9\"'“‘(\\[<\x00\x12])")
# 直後で文を区切らない略語 (末尾の . を除いて小文字)
_ABBREVIATIONS = {
    "mr", "mrs", "ms", "dr", "prof", "sr", "jr", "st", "vs", "e.g", "i.e", "cf", "al",
    "fig", "figs", "eq", "eqs", "no", "vol", "pp", "ch", "sec", "ref", "refs", "approx",
    "inc", "ltd", "co", "dept", "univ", "u.s", "a.m", "p.m",
}


def _split_at_sentence_ends(line: str) -> str:
    """英語などの文末 (. ! ?) に改行を入れる。略語・頭文字 (J. Smith)・行頭の番号 (1.) では区切らない。"""
    pieces = []
    last = 0
    for m in _SENTENCE_END.finditer(line):
        head = line[last:m.start()]
        token = head.split()[-1] if head.split() else ""
        word = token.strip("\"'“”‘’([").rstrip(".!?").lower()
        if line[m.start() - 1] == ".":
            if word in _ABBREVIATIONS or (len(word) == 1 and word.isalpha()):
                continue
            if word.isdigit() and len(head.split()) == 1:
                continue  # 行頭の番号 "1. ..."
        pieces.append(line[last:m.end(1)])
        last = m.end()
    pieces.append(line[last:])
    return "\n".join(pieces)


def _split_sentences(
    text: str, split_sentence_ends: bool = False,
) -> tuple[list[str], list[float | None], float, list[tuple[int, float]]]:
    """テキストを改行と <wait=Ns> で分割する。

    <br> は分割せず保持する（字幕でテキストボックス内改行になる）。
    {...} ブロック内の改行で分割しないよう保護する。
    <next> の位置を記録する（文の分割はしない）。

    Args:
        split_sentence_ends: True なら行の途中の文末 (. ! ?) でも区切る
            (英語のように改行なしの段落で書かれたノート向け)

    Returns:
        (sentences, pauses, leading_pause, next_positions)
        - sentences: 分割された文のリスト
        - pauses: 各文の後の無音秒数 (len = len(sentences) - 1)
          None はデフォルト pause_sec を使用、float は指定秒数
        - leading_pause: 最初の文の前の無音秒数 (0.0 = なし)
        - next_positions: [(sentence_index, char_ratio), ...]
          sentence_index=-1 は先頭 <next> (ms=0)
    """
    # {...|...} と {...} をプレースホルダに置換して分割から保護
    placeholders: list[str] = []

    def _protect(m):
        placeholders.append(m.group(0))
        return f"\x00{len(placeholders) - 1}\x00"

    protected = _READING_PATTERN.sub(_protect, text)
    protected = _BRACE_PATTERN.sub(_protect, protected)

    # <next> を抽出してプレースホルダに置換 (位置を記録するため)
    _NEXT_PH = "\x12"
    protected = _NEXT_TAG.sub(_NEXT_PH, protected)

    # 文末で区切る場合は改行を入れる ({...} と <next> はプレースホルダ化済みなので中では区切らない)
    if split_sentence_ends:
        protected = "\n".join(_split_at_sentence_ends(line) for line in protected.split("\n"))

    # 改行で分割 → 各行を <wait=Ns> でさらに分割
    sentences: list[str] = []
    pauses: list[float | None] = []
    pending_wait: float | None = None  # 次の文との間に入れるwait
    leading_pause: float = 0.0
    # <next> の位置を文ごとに記録
    next_positions: list[tuple[int, float]] = []
    pending_next: bool = False  # 文の境界に <next> がある

    lines = protected.split("\n")
    for li, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        # <wait=Ns> で分割
        parts = _WAIT_TAG.split(line)
        # parts: [text, num, unit, text, num, unit, text, ...]
        pi = 0
        while pi < len(parts):
            if pi % 3 == 0:
                # テキスト部分
                chunk = parts[pi].strip()
                if not chunk:
                    pi += 1
                    continue
                # <next> プレースホルダが含まれるか確認
                has_next = _NEXT_PH in chunk
                # <next> を除去してテキストを取得
                clean = chunk.replace(_NEXT_PH, "")
                clean = clean.strip()
                if clean:
                    if sentences:
                        pauses.append(pending_wait)
                    elif pending_wait is not None:
                        leading_pause += pending_wait
                    pending_wait = None
                    # pending_next があれば、この文の先頭に <next>
                    if pending_next:
                        if sentences:
                            # 前の文の末尾
                            next_positions.append((len(sentences) - 1, 1.0))
                        else:
                            next_positions.append((-1, 0.0))
                        pending_next = False
                    # 文中の <next> の位置を計算
                    if has_next:
                        parts_next = chunk.split(_NEXT_PH)
                        char_pos = 0
                        total_chars = len(clean)
                        for seg in parts_next[:-1]:
                            char_pos += len(seg.strip())
                            if total_chars > 0:
                                ratio = char_pos / total_chars
                            else:
                                ratio = 0.0
                            next_positions.append((len(sentences), min(ratio, 1.0)))
                    sentences.append(clean)
                elif has_next:
                    # テキストなしで <next> のみ → 境界として保留
                    pending_next = True
            elif pi % 3 == 1:
                # 数値部分 (次の pi % 3 == 2 が単位)
                num = float(parts[pi])
                unit = (parts[pi + 1] or "").lower()
                wait_sec = num / 1000 if unit == "ms" else num
                if pending_wait is None:
                    pending_wait = wait_sec
                else:
                    pending_wait += wait_sec
            # pi % 3 == 2 は単位 (pi % 3 == 1 で処理済み)
            pi += 1

    # 末尾の pending_next
    if pending_next and sentences:
        next_positions.append((len(sentences) - 1, 1.0))

    # プレースホルダを復元
    # 左から1回だけ走査する。番号順の str.replace だと、隣接するプレースホルダの
    # 境界 (例: "\x004\x002\x005\x00" の中の "\x002\x00") を誤って置換してしまう
    def _restore(s):
        return _PLACEHOLDER.sub(lambda m: placeholders[int(m.group(1))], s)

    return [_restore(s) for s in sentences], pauses, leading_pause, next_positions


# ---------------------------------------------------------------------------
# 文単位の解析結果
# ---------------------------------------------------------------------------

@dataclass
class Sentence:
    """ノートの1文の解析結果。"""
    raw: str                     # 元の文 (読み指定・タグを含む)
    display: str                 # 字幕用テキスト (書式タグ・プレースホルダを含む)
    reading: str                 # TTS に渡す読み
    speed: float | None = None   # <speed=N> (None はエンジンのデフォルト)
    pitch: float | None = None   # <pitch=N>
    intonation: float | None = None  # <intonation=N>
    volume: float | None = None  # <volume=N>
    accents: list[tuple[str, int]] = field(default_factory=list)  # {表示|読み|N} のアクセント指定


@dataclass
class ParsedNotes:
    """ノート全体の解析結果。"""
    sentences: list[Sentence]
    pauses: list[float | None]   # 各文の後の無音秒数 (None はデフォルト)
    leading_pause: float         # 最初の文の前の無音秒数
    next_positions: list[tuple[int, float]]  # <next> の位置 [(sentence_index, char_ratio), ...]


def _last_tag_value(pattern: re.Pattern, text: str) -> float | None:
    """文中で最後にマッチした音声制御タグの値を返す。"""
    matches = list(pattern.finditer(text))
    return float(matches[-1].group(1)) if matches else None


def parse_notes(text: str, split_sentence_ends: bool = False) -> ParsedNotes:
    """ノートのテキストを文単位に解析する。

    Args:
        split_sentence_ends: True なら行の途中の文末 (. ! ?) でも区切る
    """
    # <config> タグを事前除去 (configだけの行が空文にならないよう)
    text = _CONFIG_TAG.sub("", text)
    raw_sentences, pauses, leading_pause, next_positions = _split_sentences(text, split_sentence_ends)
    sentences = [
        Sentence(
            raw=s,
            display=_to_display(s),
            reading=_to_reading(s),
            speed=_last_tag_value(_SPEED_TAG, s),
            pitch=_last_tag_value(_PITCH_TAG, s),
            intonation=_last_tag_value(_INTONATION_TAG, s),
            volume=_last_tag_value(_VOLUME_TAG, s),
            accents=_extract_accents(s),
        )
        for s in raw_sentences
    ]
    return ParsedNotes(sentences, pauses, leading_pause, next_positions)
