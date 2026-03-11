"""VOICEVOX音声合成エンジン"""

import io
import re
import wave
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

from .base import TTSEngine

# 読み指定パターン: {表示テキスト|読み} or {表示テキスト|読み|アクセント位置}
_READING_PATTERN = re.compile(r"\{([^|}]+)\|([^|}]+)(?:\|(\d+))?\}")
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

# プレースホルダ: {テキスト} 内の文字をエスケープするための代替文字
_LT = "\x02"
_GT = "\x03"
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
    text = _READING_PATTERN.sub(_escape_content, text)
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


def _split_sentences(text: str) -> tuple[list[str], list[float | None], float, list[tuple[int, float]]]:
    """テキストを改行と <wait=Ns> で分割する。

    <br> は分割せず保持する（字幕でテキストボックス内改行になる）。
    {...} ブロック内の改行で分割しないよう保護する。
    <next> の位置を記録する（文の分割はしない）。

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
    def _restore(s):
        for i, orig in enumerate(placeholders):
            s = s.replace(f"\x00{i}\x00", orig)
        return s

    return [_restore(s) for s in sentences], pauses, leading_pause, next_positions


def _make_silence(params, duration_sec: float) -> bytes:
    """指定秒数の無音フレームデータを返す。"""
    num_frames = int(params.framerate * duration_sec)
    return b"\x00" * (num_frames * params.nchannels * params.sampwidth)


def _concat_wav(
    wav_chunks: list[bytes],
    pauses: list[float],
    sentences: list[str] | None = None,
    leading_pause: float = 0.0,
) -> tuple[bytes, list[tuple[str, int, int]]]:
    """複数のWAVバイナリを1つに結合する。

    Args:
        pauses: 各チャンク間の無音秒数 (len = len(wav_chunks) - 1)
        leading_pause: 最初のチャンクの前に挿入する無音秒数

    Returns:
        (結合WAV, [(文テキスト, 開始ms, 長さms), ...])
    """
    all_frames = b""
    params = None
    timings: list[tuple[str, int, int]] = []
    current_ms = 0

    for i, chunk in enumerate(wav_chunks):
        with io.BytesIO(chunk) as f:
            with wave.open(f, "rb") as w:
                if params is None:
                    params = w.getparams()
                    # 先頭の無音を挿入
                    if leading_pause > 0:
                        all_frames += _make_silence(params, leading_pause)
                        current_ms += int(leading_pause * 1000)
                frames = w.readframes(w.getnframes())
                chunk_ms = int(w.getnframes() / w.getframerate() * 1000)

        if sentences:
            timings.append((sentences[i], current_ms, chunk_ms))

        all_frames += frames
        current_ms += chunk_ms

        if i < len(pauses):
            gap = pauses[i]
            if gap > 0:
                all_frames += _make_silence(params, gap)
                current_ms += int(gap * 1000)

    if len(wav_chunks) == 1 and not pauses and leading_pause <= 0:
        return wav_chunks[0], timings

    out = io.BytesIO()
    with wave.open(out, "wb") as w:
        w.setparams(params)
        w.writeframes(all_frames)
    return out.getvalue(), timings


class VoicevoxEngine(TTSEngine):
    """VOICEVOXローカルエンジンを使った音声合成。

    事前にVOICEVOXエンジンを起動しておく必要がある。
    デフォルトで http://localhost:50021 に接続する。
    """

    def __init__(self, speaker_id: int = 1, base_url: str = "http://localhost:50021",
                 pause_sec: float = 0.5, speed_scale: float = 1.0, pitch_scale: float = 0.0,
                 intonation_scale: float = 1.0, volume_scale: float = 1.0):
        self.speaker_id = speaker_id
        self.base_url = base_url.rstrip("/")
        self.pause_sec = pause_sec
        self.speed_scale = speed_scale
        self.pitch_scale = pitch_scale
        self.intonation_scale = intonation_scale
        self.volume_scale = volume_scale

    def _audio_query(self, text: str, speed: float | None = None, pitch: float | None = None,
                     intonation: float | None = None, volume: float | None = None) -> dict:
        """テキストから音声クエリを取得する。"""
        resp = requests.post(
            f"{self.base_url}/audio_query",
            params={"text": text, "speaker": self.speaker_id},
        )
        resp.raise_for_status()
        query = resp.json()
        query["speedScale"] = speed if speed is not None else self.speed_scale
        query["pitchScale"] = pitch if pitch is not None else self.pitch_scale
        query["intonationScale"] = intonation if intonation is not None else self.intonation_scale
        query["volumeScale"] = volume if volume is not None else self.volume_scale
        return query

    def _apply_accent_overrides(self, query: dict, accents: list[tuple[str, int]]) -> dict:
        """accent_phrases のアクセント位置を上書きし、ピッチを再計算する。"""
        if not accents:
            return query
        phrases = query.get("accent_phrases", [])
        modified = False
        for katakana, accent_pos in accents:
            matched = None
            # 完全一致を優先検索
            for phrase in phrases:
                mora_text = "".join(m["text"] for m in phrase["moras"])
                if mora_text == katakana:
                    matched = phrase
                    break
            # 見つからなければ前方一致 (助詞が結合されている場合: ハシヲ vs ハシ)
            if matched is None:
                for phrase in phrases:
                    mora_text = "".join(m["text"] for m in phrase["moras"])
                    if mora_text.startswith(katakana) and len(katakana) >= 2:
                        matched = phrase
                        break
            if matched is not None:
                matched["accent"] = accent_pos
                modified = True
        if modified:
            # mora_pitch でピッチ再計算 (失敗時は mora_data を試す)
            recalculated = False
            for endpoint in ("mora_pitch", "mora_data"):
                try:
                    resp = requests.post(
                        f"{self.base_url}/{endpoint}",
                        params={"speaker": self.speaker_id},
                        json=phrases,
                    )
                    resp.raise_for_status()
                    query["accent_phrases"] = resp.json()
                    recalculated = True
                    break
                except requests.RequestException:
                    continue
            if not recalculated:
                print("[PPVoice] アクセント再計算に失敗しました (mora_pitch/mora_data 未対応)")
        return query

    def _synthesize_chunk(self, text: str) -> bytes:
        """1文のテキストからWAV音声を生成する。"""
        audio_query = self._audio_query(text)
        synth_resp = requests.post(
            f"{self.base_url}/synthesis",
            params={"speaker": self.speaker_id},
            json=audio_query,
        )
        synth_resp.raise_for_status()
        return synth_resp.content

    def _multi_synthesis(self, queries: list[dict]) -> list[bytes]:
        """複数の音声クエリを一括合成し、WAVリストを返す。"""
        resp = requests.post(
            f"{self.base_url}/multi_synthesis",
            params={"speaker": self.speaker_id},
            json=queries,
        )
        resp.raise_for_status()
        wav_list = []
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            for name in sorted(zf.namelist()):
                wav_list.append(zf.read(name))
        return wav_list

    def synthesize(self, text: str, on_chunk=None) -> bytes:
        """テキストからWAV音声を生成する。長文は文単位で分割して合成・結合する。"""
        wav, _, _ = self.synthesize_with_timings(text, on_chunk=on_chunk)
        return wav

    def synthesize_with_timings(
        self, text: str, on_chunk=None, max_workers: int = 4,
    ) -> tuple[bytes, list[tuple[str, int, int]], list[tuple[int, float]]]:
        """テキストからWAV音声を生成し、各文のタイミング情報も返す。

        audio_query を並列実行し、multi_synthesis で一括合成する。

        Args:
            on_chunk: コールバック on_chunk(chunk_index, total, sentence_text)
            max_workers: audio_query の並列数

        Returns:
            (WAVバイナリ, [(文テキスト, 開始ms, 長さms), ...],
             [(sentence_index, char_ratio), ...])
        """
        if not text:
            return b"", [], []

        # <config> タグを事前除去 (configだけの行が空文にならないよう)
        text = _CONFIG_TAG.sub("", text)

        sentences, pause_gaps, leading_pause, next_positions = _split_sentences(text)
        if not sentences:
            return b"", [], []

        # pause_gaps の None をデフォルト pause_sec に置換
        pauses = [g if g is not None else self.pause_sec for g in pause_gaps]

        display_sentences = [_to_display(s) for s in sentences]
        readings = [_to_reading(s) for s in sentences]

        # 各文の <speed>/<pitch>/<intonation>/<volume> タグを抽出 (最後にマッチした値を使用)
        speed_per_sent: list[float | None] = []
        pitch_per_sent: list[float | None] = []
        intonation_per_sent: list[float | None] = []
        volume_per_sent: list[float | None] = []
        for s in sentences:
            sm = list(_SPEED_TAG.finditer(s))
            speed_per_sent.append(float(sm[-1].group(1)) if sm else None)
            pm = list(_PITCH_TAG.finditer(s))
            pitch_per_sent.append(float(pm[-1].group(1)) if pm else None)
            im = list(_INTONATION_TAG.finditer(s))
            intonation_per_sent.append(float(im[-1].group(1)) if im else None)
            vm = list(_VOLUME_TAG.finditer(s))
            volume_per_sent.append(float(vm[-1].group(1)) if vm else None)

        # 各文のアクセント指定を抽出
        accents_per_sent = [_extract_accents(s) for s in sentences]

        # --- audio_query を並列実行 ---
        queries = [None] * len(readings)
        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            futures = {
                pool.submit(self._audio_query, r, speed_per_sent[i], pitch_per_sent[i],
                            intonation_per_sent[i], volume_per_sent[i]): i
                for i, r in enumerate(readings)
            }
            done_count = 0
            for future in as_completed(futures):
                idx = futures[future]
                queries[idx] = future.result()
                done_count += 1
                if on_chunk:
                    on_chunk(idx, len(sentences), display_sentences[idx])

        # --- アクセント上書き ---
        for i, accents in enumerate(accents_per_sent):
            if accents:
                queries[i] = self._apply_accent_overrides(queries[i], accents)

        # --- multi_synthesis で一括合成 ---
        wav_chunks = self._multi_synthesis(queries)

        wav, timings = _concat_wav(wav_chunks, pauses=pauses, sentences=display_sentences, leading_pause=leading_pause)
        return wav, timings, next_positions

    def list_speakers(self) -> list[dict]:
        """利用可能な話者一覧を取得する。"""
        resp = requests.get(f"{self.base_url}/speakers")
        resp.raise_for_status()
        return resp.json()
