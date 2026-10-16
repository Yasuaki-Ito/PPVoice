"""ノートの書き間違いの検出

知らないタグや値の形式が違うタグは、除去されずに字幕に表示され、読み上げにも含まれる。
互換性のため動作は変えず、警告だけを出す (GUI の生成前のログ・CLI の check)。
"""

import re
from dataclasses import dataclass

from i18n import t
from notes import _BRACE_PATTERN, _FORMAT_TAG, _READING_PATTERN, _SPEAKER_TAG, _STYLE_TAG, _WAIT_TAG

# タグらしいもの: < の直後が英字 (または /英字)
_TAG_LIKE = re.compile(r"</?[A-Za-z][^<>\n]*>")
_TAG_NAME = re.compile(r"</?([A-Za-z]+)")
_KNOWN_TAGS = {"b", "i", "u", "color", "font", "size", "speed", "pitch", "intonation", "volume", "br", "wait",
               "config", "cast", "dict", "next", "speaker", "style"}
# 文の中で閉じる必要のある書式タグ
_OPEN_FORMAT = re.compile(r"<(b|i|u|color|font|size)(?:=[^>]*)?>", re.IGNORECASE)
_CLOSE_FORMAT = re.compile(r"</(b|i|u|color|font|size)>", re.IGNORECASE)
# 文の区切りになるタグ (行の中で文を分ける)
_SENTENCE_BREAK = re.compile(f"{_WAIT_TAG.pattern}|{_SPEAKER_TAG.pattern}", re.IGNORECASE)


@dataclass
class Issue:
    line: int       # ノートの行 (1始まり)
    kind: str       # unknown_tag / invalid_value / unclosed_format / undefined_role / unknown_style / style_ignored
    text: str       # 問題のあるタグ
    message: str    # 利用者向けの説明
    role: str = ""  # undefined_role: 役名 / unknown_style: その時点の役
    style: str = ""  # unknown_style: スタイル名


def _mask_braces(text: str) -> str:
    """{...} の中を空白にする (改行は残すので、行番号と位置は変わらない)。"""
    def blank(m):
        return re.sub(r"[^\n]", " ", m.group(0))
    return _BRACE_PATTERN.sub(blank, _READING_PATTERN.sub(blank, text))


def lint_notes(
    text: str,
    roles: set[str] | None = None,
    default_role: str = "",
    styles: dict[str, set[str]] | None = None,
    engine: str = "voicevox",
) -> list[Issue]:
    """1枚のスライドのノートを確認する。

    Args:
        roles: 定義されている役の名前 (None なら役は確認しない)
        default_role: 既定の役の名前 (<speaker> の前の文の役)
        styles: 役名 → その話者のスタイル名 (None ならスタイルは確認しない)
        engine: voicevox / openai (OpenAI 互換では <style> を無視する)
    """
    masked = _mask_braces(text)
    issues: list[Issue] = []

    def line_of(pos: int) -> int:
        return masked.count("\n", 0, pos) + 1

    # 知らないタグ・値の形式が違うタグ
    for m in _TAG_LIKE.finditer(masked):
        tag = m.group(0)
        if _FORMAT_TAG.fullmatch(tag):
            continue
        name = _TAG_NAME.match(tag).group(1).lower()
        if name in _KNOWN_TAGS:
            issues.append(Issue(line_of(m.start()), "invalid_value", tag, t("lint_invalid_value", tag=tag)))
        else:
            issues.append(Issue(line_of(m.start()), "unknown_tag", tag, t("lint_unknown_tag", tag=tag)))

    # 文の中で閉じていない書式タグ
    for li, line in enumerate(masked.split("\n"), 1):
        for segment in _SENTENCE_BREAK.split(line):
            if not segment:
                continue
            opens = [m.group(1).lower() for m in _OPEN_FORMAT.finditer(segment)]
            closes = [m.group(1).lower() for m in _CLOSE_FORMAT.finditer(segment)]
            unclosed = [name for name in dict.fromkeys(opens + closes) if opens.count(name) != closes.count(name)]
            if unclosed:
                tags = " ".join(f"<{name}>" for name in unclosed)
                issues.append(Issue(li, "unclosed_format", tags, t("lint_unclosed", tags=tags)))

    # 役とスタイル (ノートの順に、その時点の役で確認する)
    role = default_role
    events = sorted([(m.start(), "speaker", m) for m in _SPEAKER_TAG.finditer(masked)]
                    + [(m.start(), "style", m) for m in _STYLE_TAG.finditer(masked)], key=lambda e: e[0])
    for pos, kind, m in events:
        value = m.group(1).strip()
        if kind == "speaker":
            role = value or default_role
            if roles is not None and value and value not in roles:
                issues.append(Issue(line_of(pos), "undefined_role", m.group(0),
                                    t("lint_undefined_role", name=value), role=value))
        elif engine == "openai":
            issues.append(Issue(line_of(pos), "style_ignored", m.group(0), t("lint_style_ignored")))
        elif styles is not None and role in styles and value and value not in styles[role]:
            issues.append(Issue(line_of(pos), "unknown_style", m.group(0),
                                t("lint_unknown_style", role=role, style=value), role=role, style=value))

    issues.sort(key=lambda i: i.line)
    return issues
