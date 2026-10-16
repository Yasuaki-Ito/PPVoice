"""PPTX の読み込み → 音声合成 → 埋め込み の処理 (GUI と CLI で共通)

ログは print で出す (GUI は stdout をログ欄に、CLI は stderr に流す)。
"""

import os
from dataclasses import dataclass, field
from typing import Callable

from cache import SynthCache
from i18n import t
from lint import Issue, lint_notes
from notes import _BRACE_PATTERN, _NEXT_TAG, _READING_PATTERN, Sentence, parse_notes
from pptx_reader import read_slides
from pptx_writer import _extract_click_groups, embed_audio
from settings import Settings, parse_cast_tags, parse_config_tags, parse_dict_tags
from tts.base import Timing, TTSEngine, VoiceParams
from tts.openai_compat import OpenAICompatEngine, voice_short_name
from tts.voicevox import VoicevoxEngine
from tts.wav import _concat_wav


class Cancelled(Exception):
    """生成の中断"""


# 話者一覧の取得のタイムアウト (接続, 読み込み) 秒
FETCH_TIMEOUT = (5, 30)


def _unescape_display(text: str) -> str:
    """字幕テキストのプレースホルダをログ表示用に戻す。"""
    for ph, ch in {"\x02": "<", "\x03": ">", "\x13": "$", "\x14": "$", "\x15": "{", "\x16": "}"}.items():
        text = text.replace(ph, ch)
    return text


# ---------------------------------------------------------------------------
# 設定の読み込み
# ---------------------------------------------------------------------------

def read_config(input_path: str) -> dict[str, str]:
    """PPTX のノートにある <config> タグを解析する。"""
    return parse_config_tags([s.notes_text for s in read_slides(input_path) if s.notes_text])


def read_cast(input_path: str) -> dict[str, dict[str, str]]:
    """PPTX のノートにある <cast> タグを解析する。"""
    return parse_cast_tags([s.notes_text for s in read_slides(input_path) if s.notes_text])


def read_dict(input_path: str) -> dict[str, str]:
    """PPTX のノートにある <dict> タグを解析する。"""
    return parse_dict_tags([s.notes_text for s in read_slides(input_path) if s.notes_text])


def default_role_name(settings: Settings) -> str:
    """既定の役の名前 (話者一覧を引かずに分かる範囲。話者が空なら空)。"""
    if settings.role:
        return settings.role
    if settings.engine == "openai":
        return voice_short_name(settings.speaker) if settings.speaker else ""
    return settings.speaker


# ---------------------------------------------------------------------------
# 話者一覧とエンジン
# ---------------------------------------------------------------------------

def fetch_speakers(settings: Settings):
    """話者の一覧を取得する。VOICEVOX は /speakers の結果、OpenAI互換は声の ID のリスト。"""
    engine = create_engine(settings)
    if isinstance(engine, OpenAICompatEngine):
        return engine.list_voices(timeout=FETCH_TIMEOUT)
    return engine.list_speakers(timeout=FETCH_TIMEOUT)


def resolve_voicevox_style(speakers: list[dict], speaker: str, style: str) -> tuple[int, str, str]:
    """話者名・スタイル名をスタイル ID に解決する。

    見つからない話者は一覧の最初の話者、見つからないスタイルはその話者の最初のスタイルにする
    (以前の GUI が、一覧を取得したときに最初の話者・スタイルを選んでいたのと同じ)。

    Returns:
        (スタイル ID, 実際に使う話者名, 実際に使うスタイル名)
    """
    candidates = [sp for sp in speakers if sp.get("styles")]
    if not candidates:
        raise ValueError("no speakers")
    sp = next((s for s in candidates if s["name"] == speaker), None)
    if sp is None:
        if speaker:
            print(t("log_speaker_not_found", name=speaker, used=candidates[0]["name"]))
        sp = candidates[0]
    st = next((s for s in sp["styles"] if s["name"] == style), None)
    if st is None:
        if style and sp["name"] == speaker:
            print(t("log_style_not_found", name=style, used=sp["styles"][0]["name"]))
        st = sp["styles"][0]
    return st["id"], sp["name"], st["name"]


def resolve_openai_voice(voices: list[str], voice: str) -> str:
    """声の ID を確かめる。空なら一覧の最初の声にする。一覧にない ID はそのまま使う (サーバに任せる)。"""
    if voice:
        return voice
    if not voices:
        raise ValueError("no voices")
    return voices[0]


def create_engine(settings: Settings) -> TTSEngine:
    """設定からエンジン (サーバ) を作る。声は create_cast で別に決める。"""
    if settings.engine == "openai":
        return OpenAICompatEngine(base_url=settings.url, model=settings.model.strip() or "kokoro",
                                  api_key=settings.api_key.strip())
    return VoicevoxEngine(base_url=settings.url)


@dataclass
class VoiceSpec:
    """役の声。話者・スタイルは ID に解決済み。文に音声タグがない値はここから決まる。"""
    voice: int | str          # VOICEVOX: スタイル ID / OpenAI互換: 声の ID
    speed: float = 1.0
    pitch: float = 0.0
    intonation: float = 1.0
    volume: float = 1.0
    label: str = ""           # ログ表示用 (speaker=2 / voice=af_heart, model=kokoro)
    name: str = ""            # 役名 (<speaker=...> で使う)
    speaker: str = ""         # 実際に使う話者名 (VOICEVOX のクレジット表記) / 声の ID
    styles: dict[str, int] = field(default_factory=dict)  # VOICEVOX: この話者のスタイル名 → ID (<style=...>)
    instructions: str = ""    # OpenAI互換のみ
    color: str = ""           # 字幕の文字色 (空なら字幕の既定の色)
    subtitle_name: str = ""   # 字幕に出す名前 (空なら役名)


@dataclass
class Cast:
    """役の一覧。<speaker> のない文と各スライドの先頭は既定の役で読む。

    未定義の役名・見つからないスタイルの警告は、この Cast を使う間 (1回の生成) に1度だけ出す。
    """
    default: VoiceSpec
    roles: dict[str, VoiceSpec] = field(default_factory=dict)  # <cast> の役 (名前 → 声)
    engine: str = "voicevox"
    used: dict[str, int] = field(default_factory=dict)     # 役名 → 読んだ文の数
    speakers_used: dict[str, None] = field(default_factory=dict)  # 使った話者名 (順序付きの集合)
    _warned: set = field(default_factory=set)

    @property
    def names(self) -> set[str]:
        return {self.default.name, *self.roles}

    @property
    def styles(self) -> dict[str, set[str]] | None:
        """役名 → その話者のスタイル名 (VOICEVOX のみ)。"""
        if self.engine == "openai":
            return None
        return {name: set(v.styles) for name, v in {self.default.name: self.default, **self.roles}.items()}

    def subtitle_roles(self) -> dict[str, tuple[str, str]]:
        """embed_audio に渡す {役名: (字幕色 hex RGB (空なら既定の色), 字幕に出す名前)}"""
        voices = {self.default.name: self.default, **self.roles}
        return {name: (v.color.lstrip("#"), v.subtitle_name or name) for name, v in voices.items()}

    def lint(self, text: str) -> list[Issue]:
        """ノートを確認する。役とスタイルの問題は、合成のときに同じ警告を出さないようにする。"""
        issues = lint_notes(text, self.names, self.default.name, self.styles, self.engine)
        for issue in issues:
            if issue.kind == "undefined_role":
                self._warned.add(("role", issue.role))
            elif issue.kind == "unknown_style":
                self._warned.add(("style", issue.role, issue.style))
            elif issue.kind == "style_ignored":
                self._warned.add(("style",))
        return issues

    def warn(self, key, message: str):
        if key not in self._warned:
            self._warned.add(key)
            print(message)

    def role(self, name: str | None) -> VoiceSpec:
        if name is None:
            return self.default
        if name in self.roles:
            return self.roles[name]
        if name == self.default.name:
            return self.default
        self.warn(("role", name), t("log_role_unknown", name=name, used=self.default.name))
        return self.default

    @property
    def label(self) -> str:
        """ログ表示用。役が1つなら今までどおり声だけ。"""
        if not self.roles:
            return self.default.label
        voices = {self.default.name: self.default, **self.roles}
        return ", ".join(f"{name}: {v.label}" for name, v in voices.items())


def _create_voice(settings: Settings, speakers, name: str, speaker: str, style: str, speed: float, pitch: float,
                  intonation: float, volume: float, **extra) -> VoiceSpec:
    """話者 (名前) を解決して役の声を作る。役名が空なら話者名 (OpenAI互換は声の名前の部分) にする。"""
    if settings.engine == "openai":
        voice = speaker or resolve_openai_voice(speakers, "")
        return VoiceSpec(voice, speed, pitch, intonation, volume,
                         label=f"voice={voice}, model={settings.model.strip() or 'kokoro'}",
                         name=name or voice_short_name(voice), speaker=voice, **extra)
    sid, sp_name, _ = resolve_voicevox_style(speakers, speaker, style)
    styles = {st["name"]: st["id"] for sp in speakers if sp["name"] == sp_name for st in sp.get("styles", [])}
    return VoiceSpec(sid, speed, pitch, intonation, volume, label=f"speaker={sid}",
                     name=name or sp_name, speaker=sp_name, styles=styles, **extra)


def create_cast(settings: Settings, speakers=None) -> Cast:
    """設定の既定の役と <cast> の役を解決する。speakers は取得済みの話者一覧 (None なら取得する)。"""
    need_list = settings.engine == "voicevox" or not settings.speaker or any(
        not r.speaker for r in settings.cast.values())
    if speakers is None and need_list:
        speakers = fetch_speakers(settings)
    default = _create_voice(settings, speakers, settings.role, settings.speaker, settings.style, settings.speed,
                            settings.pitch, settings.intonation, settings.volume,
                            color=settings.role_color, subtitle_name=settings.role_label)
    roles = {
        r.name: _create_voice(settings, speakers, r.name, r.speaker, r.style, r.speed, r.pitch, r.intonation,
                              r.volume, instructions=r.instructions, color=r.color, subtitle_name=r.label)
        for r in settings.cast.values()
    }
    return Cast(default, roles, engine=settings.engine)


def voice_params(sentences: list[Sentence], cast: Cast) -> list[VoiceParams]:
    """文ごとの声を決める: 文の役 (<speaker>) の声を、文の <style> と音声タグ (<speed> など) で上書きする。"""
    def pick(value, default):
        return value if value is not None else default

    result = []
    for s in sentences:
        role = cast.role(s.speaker)
        voice = role.voice
        if s.style:
            if cast.engine == "openai":
                cast.warn(("style",), t("log_style_ignored"))
            elif s.style in role.styles:
                voice = role.styles[s.style]
            else:
                cast.warn(("style", role.name, s.style),
                          t("log_style_unknown", role=role.name, speaker=role.speaker, style=s.style))
        cast.used[role.name] = cast.used.get(role.name, 0) + 1
        cast.speakers_used.setdefault(role.speaker)
        result.append(VoiceParams(voice=voice, speed=pick(s.speed, role.speed), pitch=pick(s.pitch, role.pitch),
                                  intonation=pick(s.intonation, role.intonation),
                                  volume=pick(s.volume, role.volume), instructions=role.instructions))
    return result


def synthesize_notes(
    engine: TTSEngine, text: str, cast: Cast, pause_sec: float = 0.5,
    split_sentence_ends: bool = False, on_chunk=None, turn_pause: float | None = None,
    dictionary: dict[str, str] | None = None, cache: SynthCache | None = None, on_cache_hit=None,
) -> tuple[bytes, list[Timing], list[tuple[int, float]]]:
    """ノートのテキストから WAV を作り、各文のタイミングと <next> の位置も返す。

    Args:
        pause_sec: 文と文の間の無音 (<wait> の指定がない区切り)
        split_sentence_ends: 行の途中の文末 (. ! ?) でも区切る
        on_chunk: コールバック on_chunk(文の番号, 文の数, 字幕テキスト)
        turn_pause: 役が替わる区切りの無音 (None なら pause_sec)
        dictionary: 読みの辞書 (<dict>)
        cache: 合成結果のキャッシュ (None なら使わない)。当たったら on_cache_hit() を呼ぶ

    Returns:
        (WAVバイナリ, [Timing, ...], [(sentence_index, char_ratio), ...])
    """
    if not text:
        return b"", [], []
    parsed = parse_notes(text, split_sentence_ends, dictionary)
    sentences = parsed.sentences
    if not sentences:
        return b"", [], []

    params = voice_params(sentences, cast)
    roles = [cast.role(s.speaker) for s in sentences]

    if cache is not None:
        key = cache.key(engine.cache_id(), text, split_sentence_ends, pause_sec, turn_pause, dictionary or {},
                        [vars(p) for p in params], [r.name for r in roles])
        hit = cache.get(key)
        if hit is not None:
            if on_cache_hit:
                on_cache_hit()
            return hit

    # pauses の None (<wait> の指定がない区切り) は、役が替わるなら turn_pause、それ以外は pause_sec
    def gap(i: int, g: float | None) -> float:
        if g is not None:
            return g
        if turn_pause is not None and roles[i] is not roles[i + 1]:
            return turn_pause
        return pause_sec
    pauses = [gap(i, g) for i, g in enumerate(parsed.pauses)]

    def on_done(i: int):
        if on_chunk:
            on_chunk(i, len(sentences), sentences[i].display)

    wav_chunks = engine.synthesize_sentences(sentences, params, on_done)
    wav, timings = _concat_wav(wav_chunks, pauses=pauses, sentences=[s.display for s in sentences],
                               leading_pause=parsed.leading_pause)
    for timing, role in zip(timings, roles):
        timing.speaker = role.name
    if cache is not None:
        cache.put(key, wav, timings, parsed.next_positions)
    return wav, timings, parsed.next_positions


# ---------------------------------------------------------------------------
# 生成
# ---------------------------------------------------------------------------

@dataclass
class GenerateResult:
    output: str
    roles: dict[str, int]     # 役名 → 読んだ文の数
    credits: list[str]        # VOICEVOX のクレジット表記 ("VOICEVOX:ずんだもん" など。VOICEVOX 以外は空)


def generate(
    settings: Settings,
    input_path: str,
    output_path: str,
    selected_slides: set[int] | None = None,
    engine: TTSEngine | None = None,
    cast: Cast | None = None,
    on_progress: Callable[[float], None] | None = None,
    should_cancel: Callable[[], bool] | None = None,
    cache: SynthCache | None = None,
) -> GenerateResult | None:
    """PPTX を読み込み、ノートを合成して音声付き PPTX を書き出す。

    Args:
        selected_slides: 生成するスライドの番号 (1始まり)。None なら全スライド
        engine: 作成済みのエンジン (None なら settings から作る)
        cast: 役の声 (None なら settings の話者を解決する)
        on_progress: 進捗 (0〜1) を受け取るコールバック
        should_cancel: True を返すと中断する (Cancelled を送出)
        cache: 合成結果のキャッシュ (None なら使わない)

    Returns:
        書き出した結果 (ノートのあるスライドがなければ None)
    """
    def check_cancel():
        if should_cancel and should_cancel():
            raise Cancelled()

    def progress(value):
        if on_progress:
            on_progress(value)

    if engine is None:
        engine = create_engine(settings)
    if cast is None:
        cast = create_cast(settings)

    # スライド読み込み
    print(t("log_reading_pptx", path=input_path))
    slides = read_slides(input_path)
    print(t("log_slides_found", n=len(slides)))

    # スライドフィルタ
    if selected_slides is not None:
        slides = [s for s in slides if (s.index + 1) in selected_slides]
        print(t("log_slides_selected", n=len(slides)))

    notes_count = sum(1 for s in slides if s.notes_text)
    if notes_count == 0:
        print(t("log_no_notes"))
        return None
    print(t("log_notes_count", n=notes_count))

    # ノートの書き間違い (生成は止めない)
    issues = [(info.index + 1, issue) for info in slides if info.notes_text for issue in cast.lint(info.notes_text)]
    if issues:
        print(t("log_lint_header"))
        for slide_num, issue in issues:
            print(t("log_lint_issue", slide=slide_num, line=issue.line, message=issue.message))

    # 音声合成
    print(t("log_synthesizing", desc=cast.label, pause=settings.pause))

    slide_audio = []
    slide_timings = {}
    slide_next_positions = {}
    processed = 0

    for info in slides:
        check_cancel()
        slide_num = info.index + 1
        # 進捗は処理するスライド (一部だけ選んだ場合は選んだスライド) の中での順番
        step = processed + 1
        if not info.notes_text:
            print(t("log_slide_skip", i=step, total=len(slides), num=slide_num))
            slide_audio.append((info.index, b""))
        else:
            print(t("log_slide", i=step, total=len(slides), num=slide_num))

            def on_chunk(i, total, text):
                check_cancel()
                print(f"    ({i + 1}/{total}) {_unescape_display(text)}")

            # 字幕オフでも <next> の発火タイミング計算に文のタイミングが必要なため、
            # 常にタイミング付きで合成する
            wav, timings, next_pos = synthesize_notes(engine, info.notes_text, cast, settings.pause,
                                                      settings.split_sentences, on_chunk=on_chunk,
                                                      turn_pause=settings.turn_pause,
                                                      dictionary=settings.dictionary, cache=cache,
                                                      on_cache_hit=lambda: print(t("log_cache_used")))
            slide_timings[info.index] = timings
            if next_pos:
                slide_next_positions[info.index] = next_pos
            slide_audio.append((info.index, wav))

        processed += 1
        progress(processed / len(slides))

    # PPTX出力
    print(t("log_embedding"))
    embed_audio(
        input_path,
        slide_audio,
        output_path,
        slide_timings=slide_timings,
        slide_next_positions=slide_next_positions if slide_next_positions else None,
        subtitle_roles=cast.subtitle_roles(),
        **settings.embed_kwargs(),
    )

    progress(1.0)
    if len(cast.used) > 1:
        print(t("log_role_counts", counts=", ".join(f"{name} {n}" for name, n in cast.used.items())))
    credits = []
    if isinstance(engine, VoicevoxEngine) and (engine.engine_name() or "VOICEVOX").startswith("VOICEVOX"):
        credits = [f"VOICEVOX:{name}" for name in cast.speakers_used]
        if credits:
            print(t("log_credits", credits=t("credit_sep").join(credits)))
    print(t("log_done", name=os.path.basename(output_path)))
    return GenerateResult(os.path.abspath(output_path), dict(cast.used), credits)


# ---------------------------------------------------------------------------
# <next> とアニメーションの確認
# ---------------------------------------------------------------------------

@dataclass
class NextCheck:
    slide: int          # スライド番号 (1始まり)
    animations: int     # クリックアニメーションのグループ数
    nexts: int          # <next> の数 ({…} 内は除く)

    @property
    def status(self) -> str:
        """ログ表示用の判定 (問題がなければ空)。"""
        n_clicks, n_next = self.animations, self.nexts
        if n_next > n_clicks and n_clicks > 0:
            return t("next_too_many")
        if n_next > 0 and n_clicks == 0:
            return t("next_no_anim")
        if n_clicks > n_next and n_next > 0:
            return t("next_surplus", n=n_clicks - n_next)
        if n_clicks > 0 and n_next == 0:
            return t("next_none")
        return ""


def check_next_tags(input_path: str) -> list[NextCheck]:
    """各スライドのクリックアニメーションの数と <next> の数を数える。"""
    result = []
    for si in read_slides(input_path):
        click_groups, _ = _extract_click_groups(si.slide._element)
        n_next = 0
        if si.notes_text:
            # {…} 内の <next> はエスケープ済みなので除外してカウント
            stripped = _BRACE_PATTERN.sub("", _READING_PATTERN.sub("", si.notes_text))
            n_next = len(_NEXT_TAG.findall(stripped))
        result.append(NextCheck(si.index + 1, len(click_groups), n_next))
    return result


# ---------------------------------------------------------------------------
# 字幕のプレビュー (PowerPoint で描くための一時的な PPTX)
# ---------------------------------------------------------------------------

def subtitle_roles(settings: Settings) -> dict[str, tuple[str, str]]:
    """話者一覧を引かずに分かる範囲の、役の字幕色と字幕の名前 (プレビュー用)。"""
    roles = {default_role_name(settings): (settings.role_color.lstrip("#"), settings.role_label)}
    roles.update({r.name: (r.color.lstrip("#"), r.label) for r in settings.cast.values()})
    return {name: (color, label or name) for name, (color, label) in roles.items()}


def make_preview_pptx(settings: Settings, input_path: str, slide_num: int, output_path: str,
                      sample_text: str = "") -> None:
    """slide_num (1始まり) のノートの最初の文を字幕にした PPTX を作る (音声は短い無音)。

    字幕の表示・非表示のアニメーションは画像には写らないので、字幕は1文だけにする。
    ノートがなければ sample_text を使う。
    """
    import io
    import wave

    slides = read_slides(input_path)
    info = slides[slide_num - 1]
    parsed = parse_notes(info.notes_text, settings.split_sentences, settings.dictionary) if info.notes_text else None
    if parsed and parsed.sentences:
        sentence = parsed.sentences[0]
        display, speaker = sentence.display, sentence.speaker
    else:
        display, speaker = sample_text, None
    speaker = speaker or default_role_name(settings)
    buf = io.BytesIO()
    with wave.open(buf, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(8000)
        w.writeframes(b"\x00\x00" * 800)
    embed_audio(input_path, [(info.index, buf.getvalue())], output_path,
                slide_timings={info.index: [Timing(display, 0, 100, speaker)]},
                subtitle_roles=subtitle_roles(settings), **settings.embed_kwargs())
