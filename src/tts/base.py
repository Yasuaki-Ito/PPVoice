"""TTSエンジンの抽象基底クラス"""

from abc import ABC, abstractmethod
from typing import Callable

from notes import Sentence, parse_notes

from .wav import _concat_wav


class TTSEngine(ABC):
    """音声合成エンジンの共通インターフェース

    ノートの解析 → 文ごとの合成 → 結合・タイミング計算 の流れは共通で、
    各エンジンは synthesize_sentences() だけを実装する。
    """

    def __init__(self, pause_sec: float = 0.5, split_sentence_ends: bool = False):
        self.pause_sec = pause_sec
        # 行の途中の文末 (. ! ?) でも文を区切る (英語の段落向け)
        self.split_sentence_ends = split_sentence_ends

    @abstractmethod
    def synthesize_sentences(
        self, sentences: list[Sentence], on_done: Callable[[int], None] | None = None,
    ) -> list[bytes]:
        """解析済みの各文を合成し、文と同じ順序の WAV バイナリのリストを返す。

        すべての WAV は同じフォーマット (サンプルレート・チャンネル数・サンプル幅) であること。

        Args:
            sentences: 解析済みの文 (sentence.reading を読み上げる)
            on_done: 文 i の合成 (またはその準備) が終わるたびに on_done(i) を呼ぶ (進捗表示用)
        """
        ...

    def synthesize_with_timings(
        self, text: str, on_chunk=None,
    ) -> tuple[bytes, list[tuple[str, int, int]], list[tuple[int, float]]]:
        """ノートのテキストからWAV音声を生成し、各文のタイミング情報も返す。

        Args:
            on_chunk: コールバック on_chunk(chunk_index, total, sentence_text)

        Returns:
            (WAVバイナリ, [(文テキスト, 開始ms, 長さms), ...],
             [(sentence_index, char_ratio), ...])
        """
        if not text:
            return b"", [], []

        parsed = parse_notes(text, self.split_sentence_ends)
        sentences = parsed.sentences
        if not sentences:
            return b"", [], []

        # pauses の None をデフォルト pause_sec に置換
        pauses = [g if g is not None else self.pause_sec for g in parsed.pauses]

        def on_done(i: int):
            if on_chunk:
                on_chunk(i, len(sentences), sentences[i].display)

        wav_chunks = self.synthesize_sentences(sentences, on_done)

        wav, timings = _concat_wav(
            wav_chunks, pauses=pauses, sentences=[s.display for s in sentences],
            leading_pause=parsed.leading_pause,
        )
        return wav, timings, parsed.next_positions

    def synthesize(self, text: str, on_chunk=None) -> bytes:
        """ノートのテキストからWAV音声を生成する。"""
        wav, _, _ = self.synthesize_with_timings(text, on_chunk=on_chunk)
        return wav
