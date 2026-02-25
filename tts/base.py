"""TTSエンジンの抽象基底クラス"""

from abc import ABC, abstractmethod


class TTSEngine(ABC):
    """音声合成エンジンの共通インターフェース"""

    @abstractmethod
    def synthesize(self, text: str) -> bytes:
        """テキストからWAV音声バイナリを生成する。

        Args:
            text: 読み上げるテキスト

        Returns:
            WAV形式の音声データ (bytes)
        """
        ...
