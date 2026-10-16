"""WAV の結合・無音生成"""

import io
import wave

from .base import Timing


def _make_silence(params, duration_sec: float) -> bytes:
    """指定秒数の無音フレームデータを返す。"""
    num_frames = int(params.framerate * duration_sec)
    return b"\x00" * (num_frames * params.nchannels * params.sampwidth)


def _concat_wav(
    wav_chunks: list[bytes],
    pauses: list[float],
    sentences: list[str] | None = None,
    leading_pause: float = 0.0,
) -> tuple[bytes, list[Timing]]:
    """複数のWAVバイナリを1つに結合する。

    Args:
        pauses: 各チャンク間の無音秒数 (len = len(wav_chunks) - 1)
        leading_pause: 最初のチャンクの前に挿入する無音秒数

    Returns:
        (結合WAV, [Timing(文テキスト, 開始ms, 長さms), ...])

    Raises:
        ValueError: WAV のフォーマット (サンプルレート・チャンネル数・サンプル幅) がそろっていない
    """
    all_frames = b""
    params = None
    timings: list[Timing] = []
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
                elif (w.getnchannels(), w.getsampwidth(), w.getframerate()) != \
                        (params.nchannels, params.sampwidth, params.framerate):
                    raise ValueError(
                        f"WAV format mismatch: {w.getnchannels()}ch/{w.getsampwidth() * 8}bit/{w.getframerate()}Hz "
                        f"vs {params.nchannels}ch/{params.sampwidth * 8}bit/{params.framerate}Hz")
                frames = w.readframes(w.getnframes())
                chunk_ms = int(w.getnframes() / w.getframerate() * 1000)

        if sentences:
            timings.append(Timing(sentences[i], current_ms, chunk_ms))

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
