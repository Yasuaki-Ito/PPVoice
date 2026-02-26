"""VOICEVOX音声合成エンジン"""

import io
import re
import wave
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

from .base import TTSEngine

# 読み指定パターン: {表示テキスト|読み}
_READING_PATTERN = re.compile(r"\{([^|}]+)\|([^}]+)\}")
# 保護パターン: {テキスト} (|なし) — 文分割を抑制
_BRACE_PATTERN = re.compile(r"\{([^|}]+)\}")

# 書式タグパターン: <b>, </b>, <i>, </i>, <u>, </u>, <color=#RRGGBB>, </color>,
# <font=...>, </font>, <br>, <wait=Ns>
_FORMAT_TAG = re.compile(
    r"</?(?:b|i|u|color(?:=#[0-9a-fA-F]{6})?|font(?:=[^>]+)?)>|<br\s*/?>|<wait=[\d.]+\s*(?:ms|s)?\s*>",
    re.IGNORECASE,
)

# <br> 分割パターン
_SPLIT_BR = re.compile(r"<br\s*/?>", re.IGNORECASE)

# <wait=Ns> パターン (数値+単位キャプチャ)
# 対応形式: <wait=1s>, <wait=0.5s>, <wait=500ms>, <wait=2> (単位なし=秒)
_WAIT_TAG = re.compile(r"<wait=([\d.]+)\s*(ms|s)?\s*>", re.IGNORECASE)

# プレースホルダ: {テキスト} 内の <> をエスケープするための代替文字
_LT = "\x02"
_GT = "\x03"


def _to_display(text: str) -> str:
    """{表示|読み} → 表示, {テキスト} → テキスト に変換 (字幕用)。

    <br> は改行文字に変換する。
    {テキスト} 内の <> はプレースホルダに変換し、
    書式タグとして解釈されないようにする。
    """
    text = _READING_PATTERN.sub(r"\1", text)
    # <wait> タグを除去
    text = _WAIT_TAG.sub("", text)
    # <br> → 改行
    text = _SPLIT_BR.sub("\n", text)
    # {テキスト} 内の <> をエスケープしてから展開
    def _escape_brace(m):
        return m.group(1).replace("<", _LT).replace(">", _GT)
    return _BRACE_PATTERN.sub(_escape_brace, text)


def _to_reading(text: str) -> str:
    """{表示|読み} → 読み, {テキスト} → テキスト に変換 (TTS用)。

    書式タグは読み上げに不要なため除去する。
    """
    text = _READING_PATTERN.sub(r"\2", text)
    text = _BRACE_PATTERN.sub(r"\1", text)
    return _FORMAT_TAG.sub("", text)


def _split_sentences(text: str) -> tuple[list[str], list[float | None]]:
    """テキストを改行と <wait=Ns> で分割する。

    <br> は分割せず保持する（字幕でテキストボックス内改行になる）。
    {...} ブロック内の改行で分割しないよう保護する。

    Returns:
        (sentences, pauses)
        - sentences: 分割された文のリスト
        - pauses: 各文の後の無音秒数 (len = len(sentences) - 1)
          None はデフォルト pause_sec を使用、float は指定秒数
    """
    # {...|...} と {...} をプレースホルダに置換して分割から保護
    placeholders: list[str] = []

    def _protect(m):
        placeholders.append(m.group(0))
        return f"\x00{len(placeholders) - 1}\x00"

    protected = _READING_PATTERN.sub(_protect, text)
    protected = _BRACE_PATTERN.sub(_protect, protected)

    # 改行で分割 → 各行を <wait=Ns> でさらに分割
    sentences: list[str] = []
    pauses: list[float | None] = []

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
                if chunk:
                    if sentences:
                        pauses.append(None)  # デフォルトポーズ
                    sentences.append(chunk)
            elif pi % 3 == 1:
                # 数値部分 (次の pi % 3 == 2 が単位)
                num = float(parts[pi])
                unit = (parts[pi + 1] or "").lower()
                wait_sec = num / 1000 if unit == "ms" else num
                if pauses:
                    pauses[-1] = wait_sec
            # pi % 3 == 2 は単位 (pi % 3 == 1 で処理済み)
            pi += 1
        # 改行による分割 → 次の行との間はデフォルトポーズ

    # プレースホルダを復元
    def _restore(s):
        for i, orig in enumerate(placeholders):
            s = s.replace(f"\x00{i}\x00", orig)
        return s

    return [_restore(s) for s in sentences], pauses


def _make_silence(params, duration_sec: float) -> bytes:
    """指定秒数の無音フレームデータを返す。"""
    num_frames = int(params.framerate * duration_sec)
    return b"\x00" * (num_frames * params.nchannels * params.sampwidth)


def _concat_wav(
    wav_chunks: list[bytes],
    pauses: list[float],
    sentences: list[str] | None = None,
) -> tuple[bytes, list[tuple[str, int, int]]]:
    """複数のWAVバイナリを1つに結合する。

    Args:
        pauses: 各チャンク間の無音秒数 (len = len(wav_chunks) - 1)

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

    if len(wav_chunks) == 1 and not pauses:
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

    def __init__(self, speaker_id: int = 1, base_url: str = "http://localhost:50021", pause_sec: float = 0.5):
        self.speaker_id = speaker_id
        self.base_url = base_url.rstrip("/")
        self.pause_sec = pause_sec

    def _audio_query(self, text: str) -> dict:
        """テキストから音声クエリを取得する。"""
        resp = requests.post(
            f"{self.base_url}/audio_query",
            params={"text": text, "speaker": self.speaker_id},
        )
        resp.raise_for_status()
        return resp.json()

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
        wav, _ = self.synthesize_with_timings(text, on_chunk=on_chunk)
        return wav

    def synthesize_with_timings(
        self, text: str, on_chunk=None, max_workers: int = 4,
    ) -> tuple[bytes, list[tuple[str, int, int]]]:
        """テキストからWAV音声を生成し、各文のタイミング情報も返す。

        audio_query を並列実行し、multi_synthesis で一括合成する。

        Args:
            on_chunk: コールバック on_chunk(chunk_index, total, sentence_text)
            max_workers: audio_query の並列数

        Returns:
            (WAVバイナリ, [(文テキスト, 開始ms, 長さms), ...])
        """
        if not text:
            return b"", []

        sentences, pause_gaps = _split_sentences(text)
        if not sentences:
            return b"", []

        # pause_gaps の None をデフォルト pause_sec に置換
        pauses = [g if g is not None else self.pause_sec for g in pause_gaps]

        display_sentences = [_to_display(s) for s in sentences]
        readings = [_to_reading(s) for s in sentences]

        # --- audio_query を並列実行 ---
        queries = [None] * len(readings)
        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            futures = {
                pool.submit(self._audio_query, r): i
                for i, r in enumerate(readings)
            }
            done_count = 0
            for future in as_completed(futures):
                idx = futures[future]
                queries[idx] = future.result()
                done_count += 1
                if on_chunk:
                    on_chunk(idx, len(sentences), display_sentences[idx])

        # --- multi_synthesis で一括合成 ---
        wav_chunks = self._multi_synthesis(queries)

        return _concat_wav(wav_chunks, pauses=pauses, sentences=display_sentences)

    def list_speakers(self) -> list[dict]:
        """利用可能な話者一覧を取得する。"""
        resp = requests.get(f"{self.base_url}/speakers")
        resp.raise_for_status()
        return resp.json()
