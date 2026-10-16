"""PPTX の読み込み → 音声合成 → 埋め込み の処理 (GUI と CLI で共通)

ログは print で出す (GUI は stdout をログ欄に、CLI は stderr に流す)。
"""

import os
from dataclasses import dataclass
from typing import Callable

from i18n import t
from notes import _BRACE_PATTERN, _NEXT_TAG, _READING_PATTERN
from pptx_reader import read_slides
from pptx_writer import _extract_click_groups, embed_audio
from settings import Settings, parse_config_tags
from tts.base import TTSEngine
from tts.openai_compat import OpenAICompatEngine
from tts.voicevox import VoicevoxEngine


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


# ---------------------------------------------------------------------------
# 話者一覧とエンジン
# ---------------------------------------------------------------------------

def fetch_speakers(settings: Settings):
    """話者の一覧を取得する。VOICEVOX は /speakers の結果、OpenAI互換は声の ID のリスト。"""
    if settings.engine == "openai":
        return OpenAICompatEngine(voice="", base_url=settings.url, api_key=settings.api_key).list_voices(
            timeout=FETCH_TIMEOUT)
    return VoicevoxEngine(base_url=settings.url).list_speakers(timeout=FETCH_TIMEOUT)


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


def create_engine(settings: Settings, speakers=None) -> TTSEngine:
    """設定からエンジンを作る。speakers は取得済みの話者一覧 (None なら取得する)。"""
    if settings.engine == "openai":
        voice = settings.speaker or resolve_openai_voice(
            speakers if speakers is not None else fetch_speakers(settings), "")
        return OpenAICompatEngine(
            voice=voice, base_url=settings.url, model=settings.model.strip() or "kokoro",
            api_key=settings.api_key.strip(), pause_sec=settings.pause, speed_scale=settings.speed,
            volume_scale=settings.volume, split_sentence_ends=settings.split_sentences,
        )
    if speakers is None:
        speakers = fetch_speakers(settings)
    speaker_id, _, _ = resolve_voicevox_style(speakers, settings.speaker, settings.style)
    return VoicevoxEngine(
        speaker_id=speaker_id, base_url=settings.url, pause_sec=settings.pause,
        speed_scale=settings.speed, pitch_scale=settings.pitch, intonation_scale=settings.intonation,
        volume_scale=settings.volume, split_sentence_ends=settings.split_sentences,
    )


def engine_description(settings: Settings, engine: TTSEngine) -> str:
    """ログ表示用の話者の説明。"""
    if isinstance(engine, OpenAICompatEngine):
        return f"voice={engine.voice}, model={engine.model}"
    return f"speaker={engine.speaker_id}"


# ---------------------------------------------------------------------------
# 生成
# ---------------------------------------------------------------------------

def generate(
    settings: Settings,
    input_path: str,
    output_path: str,
    selected_slides: set[int] | None = None,
    engine: TTSEngine | None = None,
    on_progress: Callable[[float], None] | None = None,
    should_cancel: Callable[[], bool] | None = None,
) -> bool:
    """PPTX を読み込み、ノートを合成して音声付き PPTX を書き出す。

    Args:
        selected_slides: 生成するスライドの番号 (1始まり)。None なら全スライド
        engine: 作成済みのエンジン (None なら settings から作る)
        on_progress: 進捗 (0〜1) を受け取るコールバック
        should_cancel: True を返すと中断する (Cancelled を送出)

    Returns:
        書き出したら True (ノートのあるスライドがなければ False)
    """
    def check_cancel():
        if should_cancel and should_cancel():
            raise Cancelled()

    def progress(value):
        if on_progress:
            on_progress(value)

    if engine is None:
        engine = create_engine(settings)

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
        return False
    print(t("log_notes_count", n=notes_count))

    # 音声合成
    print(t("log_synthesizing", desc=engine_description(settings, engine), pause=settings.pause))

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
            wav, timings, next_pos = engine.synthesize_with_timings(info.notes_text, on_chunk=on_chunk)
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
        **settings.embed_kwargs(),
    )

    progress(1.0)
    print(t("log_done", name=os.path.basename(output_path)))
    return True


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
