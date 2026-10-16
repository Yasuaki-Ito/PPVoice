"""TTSエンジンの抽象基底クラスと、合成の流れで受け渡すデータ"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Callable

from notes import Sentence


@dataclass
class VoiceParams:
    """文ごとに確定した声。役の設定と文の音声タグから pipeline が決め、エンジンはこのとおりに合成する。"""
    voice: int | str          # VOICEVOX: スタイル ID / OpenAI互換: 声の ID
    speed: float = 1.0
    pitch: float = 0.0
    intonation: float = 1.0
    volume: float = 1.0
    instructions: str = ""    # OpenAI互換のみ: 読み方の指示


@dataclass
class Timing:
    """1文の字幕の表示区間。"""
    text: str                 # 字幕テキスト (notes._to_display の出力)
    start_ms: int
    dur_ms: int
    speaker: str | None = None  # 役名


class TTSEngine(ABC):
    """音声合成エンジンの共通インターフェース

    エンジンはサーバ (URL など) だけを持ち、声は文ごとに VoiceParams で受け取る。
    ノートの解析 → 文ごとの合成 → 結合・タイミング計算の流れは pipeline.synthesize_notes にある。
    """

    def cache_id(self) -> str:
        """キャッシュのキーに入れる、エンジン (サーバ) を区別する文字列。"""
        return f"{type(self).__name__} {getattr(self, 'base_url', '')}"

    @abstractmethod
    def synthesize_sentences(
        self, sentences: list[Sentence], params: list[VoiceParams],
        on_done: Callable[[int], None] | None = None,
    ) -> list[bytes]:
        """解析済みの各文を、文ごとの声 (params[i]) で合成し、文と同じ順序の WAV バイナリのリストを返す。

        すべての WAV は同じフォーマット (サンプルレート・チャンネル数・サンプル幅) であること。

        Args:
            sentences: 解析済みの文 (sentence.reading を読み上げる)
            params: 文ごとの声 (sentences と同じ長さ)
            on_done: 文 i の合成 (またはその準備) が終わるたびに on_done(i) を呼ぶ (進捗表示用)
        """
        ...
