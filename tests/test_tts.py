"""tts/ (音声合成まわり) のテスト。サーバには接続しない。"""

import array
import io
import wave

import pytest
import requests

from tts import voicevox
from tts.openai_compat import _scale_pcm, group_voices, voice_short_name
from tts.voicevox import VoicevoxEngine, _normalize_kana
from tts.wav import _concat_wav


# ---------------------------------------------------------------------------
# アクセント指定の照合
# ---------------------------------------------------------------------------

def make_query(*phrases):
    """モーラ表記の句の並びから、audio_query の結果の形を作る (アクセントは仮に1)。"""
    return {"accent_phrases": [
        {"moras": [{"text": c} for c in p], "accent": 1, "pause_mora": None, "is_interrogative": False}
        for p in phrases
    ]}


@pytest.fixture
def engine(monkeypatch):
    """ピッチの再計算 (mora_pitch) に失敗する VOICEVOX エンジン。照合の結果だけを見る。"""
    def fail(*args, **kwargs):
        raise requests.ConnectionError("no server")
    monkeypatch.setattr(voicevox.requests, "post", fail)
    return VoicevoxEngine(speaker_id=1)


def accents(query):
    return [("".join(m["text"] for m in p["moras"]), p["accent"]) for p in query["accent_phrases"]]


def test_accent_single_word(engine):
    q = engine._apply_accent_overrides(make_query("ハシオ", "アルク"), [("ハシ", 0)])
    assert accents(q) == [("ハシオ", 0), ("アルク", 1)]


def test_accent_same_word_twice(engine):
    q = engine._apply_accent_overrides(
        make_query("ファイルオ", "トジルト", "ファイルニ"), [("ファイル", 0), ("ファイル", 0)])
    assert accents(q) == [("ファイルオ", 0), ("トジルト", 1), ("ファイルニ", 0)]


def test_accent_prefers_word_with_particle_over_compound(engine):
    q = engine._apply_accent_overrides(
        make_query("ファイルポインタワ", "ゲンゴデ", "ファイルオ"), [("ファイル", 0)])
    assert accents(q) == [("ファイルポインタワ", 1), ("ゲンゴデ", 1), ("ファイルオ", 0)]


def test_accent_same_reading_in_order(engine):
    q = engine._apply_accent_overrides(
        make_query("ハシオ", "ワタル", "ハシオ", "モツ"), [("ハシ", 2), ("ハシ", 1)])
    assert accents(q) == [("ハシオ", 2), ("ワタル", 1), ("ハシオ", 1), ("モツ", 1)]


def test_accent_merges_phrases(engine):
    q = engine._apply_accent_overrides(
        make_query("リスト", "コオゾオオ", "ツカイマス"), [(_normalize_kana("リストコウゾウ"), 4)])
    assert accents(q) == [("リストコオゾオオ", 4), ("ツカイマス", 1)]


def test_accent_long_vowel(engine):
    q = engine._apply_accent_overrides(make_query("コオゾオオ", "ツカイマス"), [("コウゾウ", 0)])
    assert accents(q) == [("コオゾオオ", 0), ("ツカイマス", 1)]


def test_accent_not_found_keeps_query(engine):
    q = engine._apply_accent_overrides(make_query("アイウ"), [("カキク", 0)])
    assert accents(q) == [("アイウ", 1)]


def test_normalize_kana():
    assert _normalize_kana("リストコウゾウ") == "リストコオゾオ"
    assert _normalize_kana("コーゾー") == "コオゾオ"
    assert _normalize_kana("ハンエイ") == "ハンエエ"
    assert _normalize_kana("ハシ") == "ハシ"


# ---------------------------------------------------------------------------
# OpenAI 互換
# ---------------------------------------------------------------------------

def test_group_voices_kokoro():
    groups = group_voices(["af_heart", "af_bella", "am_adam", "jf_alpha", "custom"])
    assert groups == [(("a", "f"), ["af_heart", "af_bella"]), (("a", "m"), ["am_adam"]),
                      (("j", "f"), ["jf_alpha"]), (None, ["custom"])]


def test_group_voices_openai_names_are_not_grouped():
    assert group_voices(["alloy", "nova", "shimmer"]) is None


def test_voice_short_name():
    assert voice_short_name("af_heart") == "heart"
    assert voice_short_name("af_amelia_inno") == "amelia_inno"
    assert voice_short_name("nova") == "nova"


def test_scale_pcm():
    pcm = array.array("h", [1000, -1000, 30000]).tobytes()
    out = array.array("h")
    out.frombytes(_scale_pcm(pcm, 2.0))
    assert list(out) == [2000, -2000, 32767]  # 範囲を超える値はクリップ
    assert _scale_pcm(pcm, 1.0) == pcm


# ---------------------------------------------------------------------------
# WAV の連結
# ---------------------------------------------------------------------------

def make_wav(ms, rate=24000):
    out = io.BytesIO()
    with wave.open(out, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(rate)
        w.writeframes(b"\x00\x00" * (rate * ms // 1000))
    return out.getvalue()


def test_concat_wav_timings():
    wav, timings = _concat_wav([make_wav(1000), make_wav(500)], pauses=[0.5], sentences=["a", "b"],
                               leading_pause=0.25)
    assert timings == [("a", 250, 1000), ("b", 1750, 500)]
    with wave.open(io.BytesIO(wav)) as w:
        assert w.getnframes() == 24000 * 2250 // 1000


def test_concat_single_wav_is_returned_as_is():
    w = make_wav(300)
    out, timings = _concat_wav([w], pauses=[], sentences=["a"])
    assert out == w
    assert timings == [("a", 0, 300)]
