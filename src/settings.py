"""PPVoice の設定 (音声・字幕・アニメーション) と <config> タグの読み書き

GUI と CLI の両方で使う。GUI はウィジェット ⇔ Settings の変換だけを行う。
話者は名前で持ち (VOICEVOX: 話者名とスタイル名 / OpenAI互換: 声の ID)、ID への解決は
エンジンを作るとき (pipeline.create_engine) に行う。
"""

import re
from dataclasses import dataclass, field, fields

DEFAULT_URLS = {"voicevox": "http://localhost:50021", "openai": "http://localhost:8880/v1"}
ENGINES = tuple(DEFAULT_URLS)

_CONFIG_RE = re.compile(r"<config\s([^>]*)>", re.IGNORECASE)
_KV_RE = re.compile(r'([\w]+)=(?:"([^"]*)"|(\S+))')

_TRUE = ("on", "true", "1")


def parse_config_tags(notes_list: list[str]) -> dict[str, str]:
    """ノートのテキストのリストから <config ...> タグを解析する。後に出てきたタグの値が優先。"""
    config: dict[str, str] = {}
    for notes in notes_list:
        for m in _CONFIG_RE.finditer(notes):
            for kv in _KV_RE.finditer(m.group(1)):
                config[kv.group(1)] = kv.group(2) if kv.group(2) is not None else kv.group(3)
    return config


def _color(value: str) -> str:
    """色の表記を #RRGGBB (大文字) にそろえる。"""
    value = value.strip()
    if not value.startswith("#"):
        value = "#" + value
    return value.upper()


@dataclass
class Settings:
    # --- エンジン ---
    engine: str = "voicevox"
    # エンジンごとの URL (切り替えても入力した URL を覚えておく)。<config> には保存しない
    urls: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_URLS))
    model: str = "kokoro"            # OpenAI互換のみ
    api_key: str = ""                # OpenAI互換のみ。PPTX に残るので <config> には保存しない
    speaker: str = ""                # VOICEVOX: 話者名 / OpenAI互換: 声の ID
    style: str = ""                  # VOICEVOX のスタイル名
    # --- 音声 ---
    speed: float = 1.0
    pitch: float = 0.0
    intonation: float = 1.0
    volume: float = 1.0
    pause: float = 0.5
    end_pause: float = 2.0
    split_sentences: bool = False    # 文末 (. ! ?) でも区切る。エンジンを切り替えると初期値に戻る
    # --- アニメーション ---
    auto_next: float = 5.0
    auto_next_enabled: bool = True
    # --- 字幕 ---
    subtitle: bool = True
    subtitle_style: str = "outline"  # outline / box
    fontsize: int = 18
    font: str = ""                   # 空ならテーマの既定のフォント
    bottom: float = 0.05
    font_color: str = "#FFFFFF"
    outline: bool = True
    outline_color: str = "#000000"
    outline_width: float = 0.75
    glow: bool = False
    glow_color: str = "#000000"
    glow_size: float = 11.0
    bg_color: str = "#000000"
    bg_alpha: int = 60
    kuten: str = "そのまま"          # 句読点置換の内部の値 (既存の <config> との互換のため日本語のまま)
    touten: str = "そのまま"
    bold: bool = False
    italic: bool = False
    underline: bool = False
    math_bold: bool = True

    @property
    def url(self) -> str:
        return self.urls.get(self.engine) or DEFAULT_URLS[self.engine]

    @url.setter
    def url(self, value: str):
        self.urls[self.engine] = value

    def set_engine(self, engine: str):
        """エンジンを切り替える。話者は選び直しになり、文末で区切るかはエンジンの初期値に戻る。"""
        if engine not in ENGINES or engine == self.engine:
            return
        self.engine = engine
        self.speaker = ""
        self.style = ""
        self.split_sentences = engine != "voicevox"

    # ------------------------------------------------------------------
    # <config> タグ
    # ------------------------------------------------------------------

    def apply_config(self, config: dict[str, str], warn=None) -> None:
        """解析済みの <config> を反映する。値の形式が違うキーは飛ばし、warn(key, value) で知らせる。

        エンジンは話者より先に切り替える (切り替えると話者が選び直しになるため)。
        """
        if "engine" in config:
            self.set_engine(config["engine"].lower())
        # 話者だけ指定してスタイルを省いた場合は、その話者の最初のスタイルにする
        if "speaker" in config and "style" not in config:
            self.style = ""
        types = {f.name: f.type for f in fields(self)}
        for key, value in config.items():
            if key in ("engine", "urls", "api_key") or key not in types:
                continue
            try:
                kind = types[key]
                if kind in ("bool", bool):
                    setattr(self, key, value.lower() in _TRUE)
                elif kind in ("int", int):
                    setattr(self, key, int(value))
                elif kind in ("float", float):
                    setattr(self, key, float(value))
                elif key.endswith("_color"):
                    setattr(self, key, _color(value))
                else:
                    setattr(self, key, value)
            except ValueError:
                if warn:
                    warn(key, value)

    def to_config_tag(self) -> str:
        """<config ...> タグの文字列にする (API キーと URL は保存しない)。"""
        parts: list[str] = []

        def add(key, val):
            s = str(val)
            parts.append(f'{key}="{s}"' if " " in s or not s else f"{key}={s}")

        def onoff(v: bool) -> str:
            return "on" if v else "off"

        add("engine", self.engine)
        if self.engine == "openai":
            add("model", self.model.strip())
            if self.speaker:
                add("speaker", self.speaker)
        else:
            if self.speaker:
                add("speaker", self.speaker)
            if self.style:
                add("style", self.style)
        # 音声
        add("pause", f"{self.pause:.1f}")
        add("speed", f"{self.speed:.1f}")
        add("pitch", f"{self.pitch:.2f}")
        add("intonation", f"{self.intonation:.1f}")
        add("volume", f"{self.volume:.1f}")
        add("end_pause", f"{self.end_pause:.1f}")
        add("split_sentences", onoff(self.split_sentences))
        add("auto_next", f"{self.auto_next:.1f}")
        add("auto_next_enabled", onoff(self.auto_next_enabled))
        # 字幕
        add("subtitle", onoff(self.subtitle))
        add("subtitle_style", self.subtitle_style)
        add("fontsize", self.fontsize)
        add("font", self.font)
        add("bottom", f"{self.bottom:.2f}")
        add("font_color", self.font_color)
        add("outline", onoff(self.outline))
        add("outline_color", self.outline_color)
        add("outline_width", f"{self.outline_width:.2f}")
        add("glow", onoff(self.glow))
        add("glow_color", self.glow_color)
        add("glow_size", f"{self.glow_size:.1f}")
        add("bg_color", self.bg_color)
        add("bg_alpha", self.bg_alpha)
        add("kuten", self.kuten)
        add("touten", self.touten)
        add("bold", onoff(self.bold))
        add("italic", onoff(self.italic))
        add("underline", onoff(self.underline))
        add("math_bold", onoff(self.math_bold))
        return "<config " + " ".join(parts) + ">"

    # ------------------------------------------------------------------
    # 字幕・埋め込みの引数
    # ------------------------------------------------------------------

    def embed_kwargs(self) -> dict:
        """pptx_writer.embed_audio に渡す字幕・タイミングの引数。"""
        return dict(
            end_pause_ms=int(self.end_pause * 1000),
            show_subtitles=self.subtitle,
            subtitle_font_size=self.fontsize,
            subtitle_font_name=self.font.strip(),
            subtitle_bottom_pct=self.bottom,
            subtitle_style=self.subtitle_style,
            subtitle_font_color=self.font_color.lstrip("#"),
            subtitle_use_outline=self.outline,
            subtitle_outline_color=self.outline_color.lstrip("#"),
            subtitle_outline_width=self.outline_width,
            subtitle_use_glow=self.glow,
            subtitle_glow_color=self.glow_color.lstrip("#"),
            subtitle_glow_size=self.glow_size,
            subtitle_bg_color=self.bg_color.lstrip("#"),
            subtitle_bg_alpha=self.bg_alpha,
            subtitle_kuten_mode=self.kuten,
            subtitle_touten_mode=self.touten,
            subtitle_default_bold=self.bold,
            subtitle_default_italic=self.italic,
            subtitle_default_underline=self.underline,
            subtitle_math_bold=self.math_bold,
            auto_next_interval_ms=int(self.auto_next * 1000) if self.auto_next_enabled else -1,
        )
