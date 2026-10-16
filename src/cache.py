"""合成結果のキャッシュ (変更したスライドだけ合成し直す)

スライドごとに、合成した WAV・文のタイミング・<next> の位置を保存する。キーは合成に使うものすべて
(ノートのテキスト・辞書・文ごとの声・間・エンジンとそのバージョン・PPVoice のバージョン) のハッシュ。
サーバ側の変更 (VOICEVOX のユーザー辞書など) はキーに入らないので、使わない設定も用意する。
"""

import hashlib
import json
import os

from tts.base import Timing

try:
    from version import __version__
except ImportError:  # テストなど、リポジトリのルートが import できないとき
    __version__ = "unknown"

LIMIT_BYTES = 500 * 1024 * 1024


def default_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "PPVoice", "cache")


class SynthCache:
    def __init__(self, directory: str | None = None, limit_bytes: int = LIMIT_BYTES):
        self.directory = directory or default_dir()
        self.limit_bytes = limit_bytes

    @staticmethod
    def key(*parts) -> str:
        data = json.dumps([__version__, *parts], ensure_ascii=False, sort_keys=True, default=str)
        return hashlib.sha256(data.encode("utf-8")).hexdigest()

    def _paths(self, key: str) -> tuple[str, str]:
        base = os.path.join(self.directory, key)
        return base + ".wav", base + ".json"

    def get(self, key: str) -> tuple[bytes, list[Timing], list[tuple[int, float]]] | None:
        wav_path, json_path = self._paths(key)
        try:
            with open(json_path, encoding="utf-8") as f:
                meta = json.load(f)
            with open(wav_path, "rb") as f:
                wav = f.read()
            # 使った日時を更新する (上限を超えたら、長く使っていないものから消す)
            for path in (wav_path, json_path):
                os.utime(path)
        except (OSError, ValueError):
            return None
        timings = [Timing(**tm) for tm in meta["timings"]]
        return wav, timings, [tuple(p) for p in meta["next_positions"]]

    def put(self, key: str, wav: bytes, timings: list[Timing], next_positions: list[tuple[int, float]]):
        wav_path, json_path = self._paths(key)
        try:
            os.makedirs(self.directory, exist_ok=True)
            with open(wav_path, "wb") as f:
                f.write(wav)
            # JSON を後に書く (JSON があれば WAV も書き終わっている)
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump({"timings": [vars(tm) for tm in timings], "next_positions": next_positions},
                          f, ensure_ascii=False)
            self._evict()
        except OSError:
            pass  # キャッシュに書けなくても生成は続ける

    def _entries(self) -> list[tuple[float, int, list[str]]]:
        """(使った日時, 大きさ, ファイル) の一覧"""
        groups: dict[str, list[str]] = {}
        for name in os.listdir(self.directory):
            stem, ext = os.path.splitext(name)
            if ext in (".wav", ".json"):
                groups.setdefault(stem, []).append(os.path.join(self.directory, name))
        entries = []
        for paths in groups.values():
            stats = [os.stat(p) for p in paths]
            entries.append((max(s.st_mtime for s in stats), sum(s.st_size for s in stats), paths))
        return entries

    def _evict(self):
        entries = sorted(self._entries())
        total = sum(size for _, size, _ in entries)
        for _, size, paths in entries:
            if total <= self.limit_bytes:
                break
            for p in paths:
                os.remove(p)
            total -= size

    def size(self) -> int:
        """キャッシュの大きさ (バイト)"""
        if not os.path.isdir(self.directory):
            return 0
        return sum(size for _, size, _ in self._entries())

    def clear(self):
        if not os.path.isdir(self.directory):
            return
        for _, _, paths in self._entries():
            for p in paths:
                os.remove(p)
