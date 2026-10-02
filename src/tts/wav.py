"""WAV の結合・無音生成"""

import io
import wave


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
