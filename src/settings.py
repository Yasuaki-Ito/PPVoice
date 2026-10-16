"""PPVoice の設定 (音声・字幕・アニメーション) と <config> / <cast> タグの読み書き

GUI と CLI の両方で使う。GUI はウィジェット ⇔ Settings の変換だけを行う。
話者は名前で持ち (VOICEVOX: 話者名とスタイル名 / OpenAI互換: 声の ID)、ID への解決は
声を作るとき (pipeline.create_cast) に行う。

役 (キャスト): <config> の話者・速度などが既定の役 (名前は role、省略すると話者名)。
<cast 名前 speaker=... ...> で役を追加し、ノートの <speaker=名前> で切り替える。
"""

import re
from dataclasses import dataclass, field, fields

DEFAULT_URLS = {"voicevox": "http://localhost:50021", "openai": "http://localhost:8880/v1"}
ENGINES = tuple(DEFAULT_URLS)

_CONFIG_RE = re.compile(r"<config\s([^>]*)>", re.IGNORECASE)
_DICT_RE = re.compile(r"<dict\s([^>]*)>", re.IGNORECASE)
_DICT_KV_RE = re.compile(r'(?:"([^"]*)"|([^\s=">]+))=(?:"([^"]*)"|([^\s>]+))')
_CAST_RE = re.compile(r'<cast\s+(?:"([^"]*)"|([^\s>="]+))([^>]*)>', re.IGNORECASE)
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


def parse_cast_tags(notes_list: list[str]) -> dict[str, dict[str, str]]:
    """ノートのテキストのリストから <cast 名前 ...> タグを解析する。同じ名前のタグは後の値が優先。"""
    casts: dict[str, dict[str, str]] = {}
    for notes in notes_list:
        for m in _CAST_RE.finditer(notes):
            name = (m.group(1) if m.group(1) is not None else m.group(2)).strip()
            if not name:
                continue
            values = casts.setdefault(name, {})
            for kv in _KV_RE.finditer(m.group(3)):
                values[kv.group(1)] = kv.group(2) if kv.group(2) is not None else kv.group(3)
    return casts


def parse_dict_tags(notes_list: list[str]) -> dict[str, str]:
    """ノートのテキストのリストから <dict 語=読み ...> タグを解析する。同じ語は後の値が優先。"""
    words: dict[str, str] = {}
    for notes in notes_list:
        for m in _DICT_RE.finditer(notes):
            for kv in _DICT_KV_RE.finditer(m.group(1)):
                word = kv.group(1) if kv.group(1) is not None else kv.group(2)
                reading = kv.group(3) if kv.group(3) is not None else kv.group(4)
                if word:
                    words[word] = reading
    return words


def _quote(key: str, value) -> str:
    s = str(value)
    return f'{key}="{s}"' if " " in s or not s else f"{key}={s}"


def _color(value: str) -> str:
    """色の表記を #RRGGBB (大文字) にそろえる。"""
    value = value.strip()
    if not value.startswith("#"):
        value = "#" + value
    return value.upper()


@dataclass
class Role:
    """役 (<cast>)。話者は名前で持つ。"""
    name: str
    speaker: str = ""                # VOICEVOX: 話者名 / OpenAI互換: 声の ID (空なら一覧の最初)
    style: str = ""                  # VOICEVOX のスタイル名 (空ならその話者の最初のスタイル)
    speed: float = 1.0
    pitch: float = 0.0
    intonation: float = 1.0
    volume: float = 1.0
    color: str = ""                  # 字幕の文字色 (空なら字幕の既定の色)
    label: str = ""                  # 字幕に出す名前 (空なら役名)
    instructions: str = ""           # OpenAI互換のみ: 読み方の指示 (gpt-4o-mini-tts など)

    def apply(self, values: dict[str, str], warn=None) -> None:
        """解析済みの <cast> の値を反映する。値の形式が違うキーは飛ばし、warn(key, value) で知らせる。"""
        for key, value in values.items():
            try:
                if key in ("speed", "pitch", "intonation", "volume"):
                    setattr(self, key, float(value))
                elif key == "color":
                    setattr(self, key, _color(value))
                elif key in ("speaker", "style", "label", "instructions"):
                    setattr(self, key, value)
                else:
                    raise ValueError(key)
            except ValueError:
                if warn:
                    warn(key, value)

    def to_tag(self) -> str:
        parts = [f'"{self.name}"' if " " in self.name else self.name]
        for key in ("speaker", "style"):
            if getattr(self, key):
                parts.append(_quote(key, getattr(self, key)))
        parts += [f"speed={self.speed:.1f}", f"pitch={self.pitch:.2f}",
                  f"intonation={self.intonation:.1f}", f"volume={self.volume:.1f}"]
        for key in ("color", "label", "instructions"):
            if getattr(self, key):
                parts.append(_quote(key, getattr(self, key)))
        return "<cast " + " ".join(parts) + ">"


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
    role: str = ""                   # 既定の役の名前 (<speaker=...> で使う)。空なら話者名
    role_color: str = ""             # 既定の役の字幕色 (空なら字幕の文字色)
    role_label: str = ""             # 既定の役の字幕に出す名前 (空なら役名)
    cast: dict[str, Role] = field(default_factory=dict)  # <cast> で追加した役 (名前 → 役)
    dictionary: dict[str, str] = field(default_factory=dict)  # <dict> の読みの辞書 (語 → 読み)
    # --- 音声 ---
    speed: float = 1.0
    pitch: float = 0.0
    intonation: float = 1.0
    volume: float = 1.0
    pause: float = 0.5
    turn_pause: float | None = None  # 話者が替わるときの間。None なら pause と同じ
    end_pause: float = 2.0
    split_sentences: bool = False    # 文末 (. ! ?) でも区切る。エンジンを切り替えると初期値に戻る
    # --- アニメーション ---
    auto_next: float = 5.0
    auto_next_enabled: bool = True
    # --- 字幕 ---
    subtitle: bool = True
    subtitle_style: str = "outline"  # outline / box
    subtitle_name: bool = False      # 字幕の前に話者名を付ける (掛け合い向け)
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
            if key in ("engine", "urls", "api_key", "cast", "dictionary") or key not in types:
                continue
            try:
                kind = types[key]
                if key == "turn_pause":
                    self.turn_pause = float(value) if value.strip() else None
                elif kind in ("bool", bool):
                    setattr(self, key, value.lower() in _TRUE)
                elif kind in ("int", int):
                    setattr(self, key, int(value))
                elif kind in ("float", float):
                    setattr(self, key, float(value))
                elif key.endswith("_color"):
                    setattr(self, key, _color(value) if value.strip() else "")
                else:
                    setattr(self, key, value)
            except ValueError:
                if warn:
                    warn(key, value)

    def apply_cast(self, casts: dict[str, dict[str, str]], warn=None) -> None:
        """解析済みの <cast> を役に反映する。既にある役は、書かれたキーだけ上書きする。"""
        for name, values in casts.items():
            self.cast.setdefault(name, Role(name)).apply(values, warn=warn)

    def to_cast_tags(self) -> list[str]:
        """<cast ...> タグの文字列のリスト (役の数だけ)。"""
        return [role.to_tag() for role in self.cast.values()]

    def to_dict_tag(self) -> str:
        """<dict ...> タグの文字列 (辞書が空なら空文字列)。"""
        def q(s: str) -> str:
            return f'"{s}"' if not s or any(c in s for c in ' ="') else s
        if not self.dictionary:
            return ""
        return "<dict " + " ".join(f"{q(w)}={q(r)}" for w, r in self.dictionary.items()) + ">"

    def to_config_tag(self) -> str:
        """<config ...> タグの文字列にする (API キーと URL は保存しない)。"""
        parts: list[str] = []

        def add(key, val):
            parts.append(_quote(key, val))

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
        if self.role:
            add("role", self.role)
        if self.role_color:
            add("role_color", self.role_color)
        if self.role_label:
            add("role_label", self.role_label)
        # 音声
        add("pause", f"{self.pause:.1f}")
        if self.turn_pause is not None:
            add("turn_pause", f"{self.turn_pause:.1f}")
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
        if self.subtitle_name:
            add("subtitle_name", onoff(self.subtitle_name))
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
            subtitle_show_name=self.subtitle_name,
            auto_next_interval_ms=int(self.auto_next * 1000) if self.auto_next_enabled else -1,
        )
