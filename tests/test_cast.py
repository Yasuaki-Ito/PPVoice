"""複数話者 (<speaker> / <style> / <cast>) のテスト。サーバには接続しない。"""

import io
import wave

from notes import parse_notes
from pipeline import Cast, VoiceSpec, create_cast, default_role_name, synthesize_notes, voice_params
from settings import Role, Settings, parse_cast_tags
from tts.base import TTSEngine, VoiceParams

SPEAKERS = [
    {"name": "四国めたん", "styles": [{"name": "ノーマル", "id": 2}, {"name": "あまあま", "id": 0}]},
    {"name": "ずんだもん", "styles": [{"name": "ノーマル", "id": 3}, {"name": "ささやき", "id": 22}]},
]


def speakers_and_styles(text, split=False):
    return [(s.display, s.speaker, s.style) for s in parse_notes(text, split).sentences]


# ---------------------------------------------------------------------------
# ノートの解析
# ---------------------------------------------------------------------------

def test_speaker_continues_until_next_speaker():
    assert speakers_and_styles("一。\n<speaker=A>二。\n三。\n<speaker=B>四。") == [
        ("一。", None, None), ("二。", "A", None), ("三。", "A", None), ("四。", "B", None)]


def test_speaker_in_the_middle_of_a_line_splits_the_sentence():
    assert speakers_and_styles("<speaker=A>一。<speaker=B>二。<next>三。") == [
        ("一。", "A", None), ("二。三。", "B", None)]


def test_speaker_only_line_applies_to_following_sentences():
    p = parse_notes("一。\n<speaker=B>\n二。")
    assert [(s.display, s.speaker) for s in p.sentences] == [("一。", None), ("二。", "B")]
    assert p.pauses == [None]  # タグだけの行で区切りが増えない


def test_speaker_name_with_quotes_and_japanese():
    assert speakers_and_styles('<speaker="ずんだもん">一。\n<speaker= めたん >二。') == [
        ("一。", "ずんだもん", None), ("二。", "めたん", None)]


def test_style_applies_to_its_sentence_only():
    assert speakers_and_styles("<speaker=A><style=ささやき>一。\n二。") == [
        ("一。", "A", "ささやき"), ("二。", "A", None)]


def test_speaker_and_style_are_removed_from_reading():
    s = parse_notes("<speaker=A><style=ささやき>こんにちは。").sentences[0]
    assert s.reading == "こんにちは。"


def test_speaker_inside_braces_is_text():
    s = parse_notes("{<speaker=A>} と書く。").sentences[0]
    assert s.speaker is None
    assert s.display == "\x02speaker=A\x03 と書く。"


def test_speaker_with_sentence_ends():
    assert speakers_and_styles("<speaker=A>Hello there. <speaker=B>Hi. How are you?", split=True) == [
        ("Hello there.", "A", None), ("Hi.", "B", None), ("How are you?", "B", None)]


def test_cast_tag_is_not_read():
    p = parse_notes('<cast A speaker="ずんだもん" color=#A6E22E>\n一。')
    assert [s.display for s in p.sentences] == ["一。"]


def test_each_slide_starts_with_default_role():
    # スライドごとに parse_notes するので、前のスライドの <speaker> は持ち越さない
    parse_notes("<speaker=B>前のスライド。")
    assert parse_notes("次のスライド。").sentences[0].speaker is None


# ---------------------------------------------------------------------------
# <cast> タグ
# ---------------------------------------------------------------------------

def test_parse_cast_tags():
    casts = parse_cast_tags([
        '<cast A speaker="ずんだもん" style=ノーマル color=#a6e22e>',
        '<cast "B 2" speaker=四国めたん speed=1.1 label="めたん先生">\n<CAST A speed=1.2>',
    ])
    assert casts == {
        "A": {"speaker": "ずんだもん", "style": "ノーマル", "color": "#a6e22e", "speed": "1.2"},
        "B 2": {"speaker": "四国めたん", "speed": "1.1", "label": "めたん先生"},
    }


def test_apply_cast_and_roundtrip():
    s = Settings()
    warned = []
    s.apply_cast(parse_cast_tags(['<cast A speaker=ずんだもん color=#a6e22e speed=abc foo=1>']),
                 warn=lambda k, v: warned.append(k))
    assert s.cast == {"A": Role("A", speaker="ずんだもん", color="#A6E22E")}
    assert warned == ["speed", "foo"]
    s.cast["B"] = Role("B", speaker="四国めたん", style="あまあま", speed=1.1, label="め たん",
                       instructions="Speak calmly.")
    s2 = Settings()
    s2.apply_cast(parse_cast_tags(s.to_cast_tags()))
    assert s2.cast == s.cast


def test_role_and_turn_pause_in_config():
    s = Settings()
    s.apply_config({"role": "ナレーター", "turn_pause": "0.2"})
    assert (s.role, s.turn_pause) == ("ナレーター", 0.2)
    tag = s.to_config_tag()
    assert "role=ナレーター" in tag and "turn_pause=0.2" in tag
    assert "role=" not in Settings().to_config_tag() and "turn_pause" not in Settings().to_config_tag()


def test_default_role_name():
    assert default_role_name(Settings(speaker="ずんだもん")) == "ずんだもん"
    assert default_role_name(Settings(speaker="ずんだもん", role="A")) == "A"
    assert default_role_name(Settings(engine="openai", speaker="af_heart")) == "heart"


# ---------------------------------------------------------------------------
# 役の解決と文ごとの声
# ---------------------------------------------------------------------------

def make_settings():
    s = Settings(speaker="ずんだもん", speed=1.1)
    s.cast = {"めたん": Role("めたん", speaker="四国めたん", style="あまあま", pitch=0.05)}
    return s


def test_create_cast():
    cast = create_cast(make_settings(), speakers=SPEAKERS)
    assert (cast.default.name, cast.default.voice, cast.default.speed) == ("ずんだもん", 3, 1.1)
    metan = cast.roles["めたん"]
    assert (metan.voice, metan.speaker, metan.pitch) == (0, "四国めたん", 0.05)
    assert metan.styles == {"ノーマル": 2, "あまあま": 0}
    assert cast.label == "ずんだもん: speaker=3, めたん: speaker=0"


def test_voice_params_per_role_and_style(capsys):
    cast = create_cast(make_settings(), speakers=SPEAKERS)
    sentences = parse_notes("一。\n<speaker=めたん>二。\n<style=ノーマル><speed=1.5>三。\n"
                            "<speaker=ずんだもん><style=ささやき>四。\n<style=ない>五。\n<speaker=誰>六。\n"
                            "<speaker=誰>七。").sentences
    assert voice_params(sentences, cast) == [
        VoiceParams(3, 1.1, 0.0, 1.0, 1.0),
        VoiceParams(0, 1.0, 0.05, 1.0, 1.0),
        VoiceParams(2, 1.5, 0.05, 1.0, 1.0),
        VoiceParams(22, 1.1, 0.0, 1.0, 1.0),
        VoiceParams(3, 1.1, 0.0, 1.0, 1.0),
        VoiceParams(3, 1.1, 0.0, 1.0, 1.0),
        VoiceParams(3, 1.1, 0.0, 1.0, 1.0),
    ]
    out = capsys.readouterr().out
    assert out.count("誰") == 1  # 未定義の役の警告は1度だけ
    assert "ない" in out
    assert cast.used == {"ずんだもん": 5, "めたん": 2}
    assert list(cast.speakers_used) == ["ずんだもん", "四国めたん"]


def test_openai_cast_ignores_style(capsys):
    s = Settings(engine="openai", speaker="af_heart")
    s.cast = {"adam": Role("adam", speaker="am_adam", instructions="Calm.")}
    cast = create_cast(s)
    sentences = parse_notes("<style=x>Hi.\n<speaker=adam>Hello.").sentences
    assert voice_params(sentences, cast) == [VoiceParams("af_heart"), VoiceParams("am_adam", instructions="Calm.")]
    assert "<style>" in capsys.readouterr().out


class FakeEngine(TTSEngine):
    """1文を 100ms の無音にするエンジン"""

    def synthesize_sentences(self, sentences, params, on_done=None):
        out = io.BytesIO()
        with wave.open(out, "wb") as w:
            w.setnchannels(1)
            w.setsampwidth(2)
            w.setframerate(1000)
            w.writeframes(b"\x00\x00" * 100)
        return [out.getvalue()] * len(sentences)


def test_turn_pause_and_timing_speaker():
    cast = Cast(VoiceSpec(1, name="A"), {"B": VoiceSpec(2, name="B")})
    text = "一。\n二。\n<speaker=B>三。\n<speaker=A>四。<wait=1s>\n<speaker=B>五。"
    _, timings, _ = synthesize_notes(FakeEngine(), text, cast, pause_sec=0.5, turn_pause=0.2)
    assert [(t.speaker, t.start_ms) for t in timings] == [
        ("A", 0), ("A", 600), ("B", 900), ("A", 1200), ("B", 2300)]
    # turn_pause を指定しなければ pause と同じ
    _, timings, _ = synthesize_notes(FakeEngine(), text, cast, pause_sec=0.5)
    assert [t.start_ms for t in timings] == [0, 600, 1200, 1800, 2900]


def test_three_roles_alternating():
    cast = Cast(VoiceSpec(1, name="A"), {"B": VoiceSpec(2, name="B"), "C": VoiceSpec(3, name="C")})
    sentences = parse_notes("<speaker=A>一。\n<speaker=B>二。\n<speaker=C>三。\n<speaker=A>四。").sentences
    assert [p.voice for p in voice_params(sentences, cast)] == [1, 2, 3, 1]


def test_voicevox_accent_uses_style_of_the_sentence(monkeypatch):
    # <style> で替えたスタイルの ID で、アクセント指定のピッチを計算し直す (mora_pitch)
    from tts import voicevox
    engine = voicevox.VoicevoxEngine()
    calls = []

    class Resp:
        def __init__(self, data):
            self.data = data

        def raise_for_status(self):
            pass

        def json(self):
            return self.data

    def post(url, params=None, json=None, **kw):
        calls.append((url.rsplit("/", 1)[-1], params["speaker"]))
        return Resp(json)

    monkeypatch.setattr(voicevox.requests, "post", post)
    monkeypatch.setattr(engine, "_audio_query", lambda text, p: {"accent_phrases": [
        {"moras": [{"text": "ハ"}, {"text": "シ"}], "accent": 1, "pause_mora": None}]})
    monkeypatch.setattr(engine, "_multi_synthesis", lambda queries, sid: [b""] * len(queries))
    cast = create_cast(make_settings(), speakers=SPEAKERS)
    sentences = parse_notes("<style=ささやき>{端|はし|0}。\n<speaker=めたん>{端|はし|0}。").sentences
    engine.synthesize_sentences(sentences, voice_params(sentences, cast))
    assert calls == [("mora_pitch", 22), ("mora_pitch", 0)]
