"""合成結果のキャッシュ (cache.py) のテスト"""

import os

from cache import SynthCache
from pipeline import Cast, VoiceSpec, synthesize_notes
from test_cast import FakeEngine
from tts.base import Timing


class CountingEngine(FakeEngine):
    def __init__(self):
        self.calls = 0

    def synthesize_sentences(self, sentences, params, on_done=None):
        self.calls += 1
        return super().synthesize_sentences(sentences, params, on_done)


def test_put_and_get(tmp_path):
    cache = SynthCache(str(tmp_path))
    key = cache.key("a", 1)
    assert cache.get(key) is None
    cache.put(key, b"wav", [Timing("文", 0, 100, "A")], [(0, 0.5)])
    assert cache.get(key) == (b"wav", [Timing("文", 0, 100, "A")], [(0, 0.5)])
    assert cache.key("a", 1) == key and cache.key("a", 2) != key


def test_evicts_least_recently_used(tmp_path):
    cache = SynthCache(str(tmp_path), limit_bytes=2500)
    for i, name in enumerate(["old", "mid", "new"]):
        cache.put(name, b"x" * 1000, [], [])
        for ext in (".wav", ".json"):
            os.utime(tmp_path / (name + ext), (1000 + i, 1000 + i))
    cache.put("newest", b"x" * 1000, [], [])
    assert cache.get("old") is None and cache.get("mid") is None
    assert cache.get("new") is not None and cache.get("newest") is not None


def test_synthesize_notes_uses_cache(tmp_path):
    cache = SynthCache(str(tmp_path))
    engine = CountingEngine()
    cast = Cast(VoiceSpec(1, name="A"))
    hits = []
    first = synthesize_notes(engine, "一。\n二。", cast, cache=cache)
    second = synthesize_notes(engine, "一。\n二。", cast, cache=cache, on_cache_hit=lambda: hits.append(1))
    assert first == second and engine.calls == 1 and hits == [1]
    # 声・間・テキストが変われば合成し直す
    synthesize_notes(engine, "一。\n二。", Cast(VoiceSpec(1, speed=1.2, name="A")), cache=cache)
    synthesize_notes(engine, "一。\n二。", cast, pause_sec=0.3, cache=cache)
    synthesize_notes(engine, "一。\n三。", cast, cache=cache)
    assert engine.calls == 4


def test_clear(tmp_path):
    cache = SynthCache(str(tmp_path))
    cache.put("k", b"wav", [], [])
    assert cache.size() > 0
    cache.clear()
    assert cache.size() == 0 and cache.get("k") is None
