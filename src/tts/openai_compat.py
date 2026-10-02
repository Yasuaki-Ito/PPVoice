"""OpenAI 互換 TTS エンジン (Kokoro-FastAPI, OpenAI など)

POST {base_url}/audio/speech に1文ずつ問い合わせる。
音声は response_format="pcm" (24kHz / 16bit / モノラルの生データ) で受け取り、WAV に包む。
WAV で受け取るとストリーミング時にヘッダの長さが不正確な場合があるため。
"""

import array
import io
import sys
import wave
from concurrent.futures import ThreadPoolExecutor

import requests

from notes import Sentence

from .base import TTSEngine

# OpenAI の pcm 形式 (Kokoro-FastAPI も同じ)
PCM_SAMPLE_RATE = 24000
PCM_SAMPLE_WIDTH = 2
PCM_CHANNELS = 1

# /audio/voices が使えないサーバ (OpenAI 本家など) 向けの既定の声
DEFAULT_OPENAI_VOICES = ["alloy", "ash", "ballad", "coral", "echo", "fable", "nova", "onyx", "sage", "shimmer"]


def _pcm_to_wav(pcm: bytes) -> bytes:
    out = io.BytesIO()
    with wave.open(out, "wb") as w:
        w.setnchannels(PCM_CHANNELS)
        w.setsampwidth(PCM_SAMPLE_WIDTH)
        w.setframerate(PCM_SAMPLE_RATE)
        w.writeframes(pcm)
    return out.getvalue()


def _scale_pcm(pcm: bytes, factor: float) -> bytes:
    """16bit PCM の音量を factor 倍にする (クリップあり)。"""
    if factor == 1.0:
        return pcm
    samples = array.array("h")
    samples.frombytes(pcm[: len(pcm) // 2 * 2])
    if sys.byteorder != "little":
        samples.byteswap()
    samples = array.array("h", (max(-32768, min(32767, int(s * factor))) for s in samples))
    if sys.byteorder != "little":
        samples.byteswap()
    return samples.tobytes()


class OpenAICompatEngine(TTSEngine):
    """OpenAI 互換の音声合成 API を使うエンジン。

    base_url は OpenAI SDK と同じく /v1 まで含める (例: http://localhost:8880/v1)。
    """

    # VOICEVOX 専用で、このエンジンでは使えない機能
    UNSUPPORTED = ("pitch", "intonation", "accent")

    def __init__(self, voice: str, base_url: str = "http://localhost:8880/v1",
                 model: str = "kokoro", api_key: str = "",
                 pause_sec: float = 0.5, speed_scale: float = 1.0, volume_scale: float = 1.0,
                 max_workers: int = 2, timeout: float = 120, split_sentence_ends: bool = True):
        super().__init__(pause_sec=pause_sec, split_sentence_ends=split_sentence_ends)
        self.voice = voice
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.api_key = api_key
        self.speed_scale = speed_scale
        self.volume_scale = volume_scale
        self.max_workers = max_workers
        self.timeout = timeout

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self.api_key}"} if self.api_key else {}

    def _speech(self, s: Sentence) -> bytes:
        """1文を合成して WAV を返す。"""
        speed = s.speed if s.speed is not None else self.speed_scale
        resp = requests.post(
            f"{self.base_url}/audio/speech",
            headers=self._headers(),
            json={
                "model": self.model,
                "input": s.reading,
                "voice": self.voice,
                "response_format": "pcm",
                "speed": speed,
                "stream": False,
            },
            timeout=self.timeout,
        )
        resp.raise_for_status()
        volume = s.volume if s.volume is not None else self.volume_scale
        return _pcm_to_wav(_scale_pcm(resp.content, volume))

    def synthesize_sentences(self, sentences: list[Sentence], on_done=None) -> list[bytes]:
        self._warn_unsupported(sentences)
        with ThreadPoolExecutor(max_workers=self.max_workers) as pool:
            futures = [pool.submit(self._speech, s) for s in sentences]
            wavs = []
            for i, future in enumerate(futures):
                wavs.append(future.result())
                if on_done:
                    on_done(i)
        return wavs

    @staticmethod
    def _warn_unsupported(sentences: list[Sentence]):
        used = []
        if any(s.pitch is not None for s in sentences):
            used.append("<pitch>")
        if any(s.intonation is not None for s in sentences):
            used.append("<intonation>")
        if any(s.accents for s in sentences):
            used.append("アクセント指定 {…|…|N}")
        if used:
            print(f"[PPVoice] このエンジンでは {', '.join(used)} は使えないため無視します")

    def list_voices(self, timeout: float | tuple[float, float] | None = None) -> list[str]:
        """利用できる声の一覧を返す。

        /audio/voices (Kokoro-FastAPI など) が使えなければ OpenAI の既定の声を返す。
        """
        try:
            resp = requests.get(f"{self.base_url}/audio/voices", headers=self._headers(), timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
        except requests.exceptions.ConnectionError:
            raise
        except (requests.RequestException, ValueError):
            return list(DEFAULT_OPENAI_VOICES)
        voices = data.get("voices", data) if isinstance(data, dict) else data
        # Kokoro-FastAPI: [{"id": ..., "name": ...}] (旧形式は文字列のリスト)
        return [v["id"] if isinstance(v, dict) else str(v) for v in voices]
