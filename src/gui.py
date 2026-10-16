"""PowerPoint自動スピーチツール GUI (customtkinter)"""

import os
import re
import sys
import threading
import webbrowser
import tkinter as tk
import winsound
from tkinter import colorchooser, filedialog, font as tkfont, messagebox

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import customtkinter as ctk
import requests
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except ImportError:
    _HAS_DND = False

import pipeline
from pptx_reader import read_slides
from settings import DEFAULT_URLS, ENGINES, Settings
from tts.openai_compat import KOKORO_LANG_CODES, group_voices, voice_short_name
from version import __version__
import i18n
from i18n import t

# 句読点の置換先: (内部の値, 表示用の文言キー)。内部の値は <config> の kuten/touten に保存する値で、
# 既存の PPTX との互換のため日本語の表記のまま。表示用の文言キーが None なら値をそのまま表示する
_TOUTEN_CHOICES = [("そのまま", "punct_unchanged"), ("、", None), (",(半角)", "punct_comma_half"),
                   ("，(全角)", "punct_comma_full"), ("(半角空白)", "punct_space_half"), ("(全角空白)", "punct_space_full")]
_KUTEN_CHOICES = [("そのまま", "punct_unchanged"), ("。", None), (".(半角)", "punct_period_half"),
                  ("．(全角)", "punct_period_full"), ("(半角空白)", "punct_space_half"), ("(全角空白)", "punct_space_full")]

# ドキュメントサイト・リポジトリ
DOCS_URL = "https://yasuaki-ito.github.io/PPVoice/"
DOCS_URL_EN = "https://yasuaki-ito.github.io/PPVoice/en/"
GITHUB_URL = "https://github.com/Yasuaki-Ito/PPVoice"

# 字幕テキスト中のプレースホルダ → ログ表示用の文字
_DISPLAY_UNESCAPE = {"\x02": "<", "\x03": ">", "\x13": "$", "\x14": "$", "\x15": "{", "\x16": "}"}


def _unescape_display(text: str) -> str:
    """字幕テキストのプレースホルダをログ表示用に戻す。"""
    for ph, ch in _DISPLAY_UNESCAPE.items():
        text = text.replace(ph, ch)
    return text

ctk.set_appearance_mode("light")
ctk.set_default_color_theme(
    os.path.join(os.path.dirname(__file__), "theme_modern.json")
)


# 生成処理の中断を伝える例外
_CancelledError = pipeline.Cancelled


class LogRedirector:
    """print 出力を CTkTextbox にリダイレクトする。"""

    def __init__(self, textbox: ctk.CTkTextbox):
        self.textbox = textbox

    def write(self, text):
        self.textbox.after(0, self._append, text)

    def _append(self, text):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", text)
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def flush(self):
        pass


if _HAS_DND:
    class _AppBase(ctk.CTk, TkinterDnD.DnDWrapper):
        def __init__(self):
            super().__init__()
            try:
                self.TkdndVersion = TkinterDnD._require(self)
            except Exception:
                self.TkdndVersion = None
else:
    _AppBase = ctk.CTk


class App(_AppBase):
    # TTS エンジンの内部キー
    _ENGINES = ENGINES
    _DEFAULT_URLS = DEFAULT_URLS

    def __init__(self):
        super().__init__()
        self.title(f"PPVoice v{__version__}")
        self.geometry("1050x650")
        self.minsize(960, 500)

        # アイコン設定 (src/ 内 → ルート の順で探す)
        for _d in [os.path.dirname(__file__), os.path.join(os.path.dirname(__file__), "..")]:
            ico_path = os.path.join(_d, "app.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
                break

        self._speakers_cache: list[dict] = []
        # 話者 (VOICEVOX: スタイルラベル → スタイルID, OpenAI互換: 声の名前 → 声の名前)
        self._speaker_map: dict[str, int | str] = {}
        self._engine = "voicevox"
        # エンジンごとの URL (切り替え時に入力中の値を覚えておく)
        self._engine_urls = dict(self._DEFAULT_URLS)
        # 表示言語 (保存された設定 → OS の表示言語)
        i18n.load_language()
        self._styles_by_speaker: dict[str, list[tuple[str, int]]] = {}
        # OpenAI互換の声の一覧 (取得した順) と、Kokoro の命名規則で2段階に分けたか
        self._openai_voices: list[str] = []
        self._voices_grouped = False
        self._running = False
        self._cancel_event = threading.Event()
        # 選びたい話者 (話者名, スタイル名)。OpenAI互換は (声の ID, "")。
        # 話者一覧を取得したときに、この名前の話者を選ぶ
        self._wanted_voice: tuple[str, str] = ("", "")
        # 画面にまだ欄のない設定 (役・話者交代の間)。<config> / <cast> から読み、そのまま保存・生成に使う
        self._voice_extra = Settings()
        self._test_stop = False

        self._build_ui()
        self._setup_dnd()
        self._update_run_btn()

    def _setup_dnd(self):
        """ドラッグ&ドロップを設定する (tkinterdnd2)。"""
        if not _HAS_DND or not getattr(self, "TkdndVersion", None):
            return
        try:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_file_drop)
        except Exception:
            pass

    def _on_file_drop(self, event):
        """ファイルがドロップされた時の処理。"""
        try:
            files = self.tk.splitlist(event.data)
        except Exception:
            files = [event.data]
        for f in files:
            f = f.strip("{}")
            if f.lower().endswith(".pptx"):
                self._set_input_file(f)
                break

    # ------------------------------------------------------------------
    # UI構築
    # ------------------------------------------------------------------

    def _build_ui(self):
        # 2カラムレイアウト: 左=設定, 右=生成・ログ
        # (言語の切り替え時は container ごと作り直す → _rebuild_ui)
        self._container = container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=16, pady=16)
        container.grid_columnconfigure(0, weight=3, minsize=560)
        container.grid_columnconfigure(1, weight=2, minsize=280)
        container.grid_rowconfigure(0, weight=1)

        # 左カラム (スクロール可能な設定パネル)
        self._left = left = ctk.CTkScrollableFrame(container, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self._build_file_section(left)
        self._build_voice_section(left)
        self._build_subtitle_section(left)
        self._build_animation_section(left)

        # 設定保存ボタン (左カラム最下部)
        ctk.CTkButton(
            left, text=t("save_config"), width=100, height=30, command=self._on_save_config,
        ).pack(anchor="e", padx=14, pady=(8, 4))

        # 右カラム (生成・ログ)
        right = ctk.CTkFrame(container, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self._build_action_section(right)
        self.input_var.trace_add("write", lambda *_: self._update_run_btn())

    def _on_language_selected(self, label: str):
        lang = next(k for k, v in i18n.LANGUAGES.items() if v == label)
        if lang == i18n.get_language():
            return
        if self._running:
            # 生成中は作り直せないので元に戻す
            self.language_selector.set(i18n.LANGUAGES[i18n.get_language()])
            return
        i18n.set_language(lang)
        i18n.save_language(lang)
        self._rebuild_ui()

    def _rebuild_ui(self):
        """画面を現在の言語で作り直す。入力中の設定・取得済みの話者一覧・ログは引き継ぐ。"""
        settings = self._settings_from_ui()
        state = {
            "input": self.input_var.get(), "output": self.output_var.get(),
            "slide_range": self.slide_range_var.get(), "selected": self._selected_slides,
            "test_text": self.test_textbox.get("1.0", "end-1c"),
            "log": self.log_box.get("1.0", "end-1c"),
        }
        self._container.destroy()
        self._build_ui()

        # 入力・出力ファイル (_set_input_file は <config> を読み直すので使わない)
        self.input_var.set(state["input"])
        self.output_var.set(state["output"])
        self.slide_range_var.set(state["slide_range"])
        self._on_slide_range_changed()
        self._selected_slides = state["selected"]
        # エンジンと取得済みの話者一覧
        self._apply_engine_ui()
        if self._engine == "voicevox" and self._speakers_cache:
            self._on_speakers_fetched(self._speakers_cache, quiet=True)
        elif self._engine == "openai" and self._openai_voices:
            self._on_voices_fetched(self._engine, self._openai_voices, quiet=True)
        # 話者の選択・音声・字幕などの設定
        self._load_settings_to_ui(settings)
        # テスト欄が初期値のままなら、新しい言語の初期値にする
        if state["test_text"] not in i18n.STRINGS["test_text"]:
            self.test_textbox.delete("1.0", "end")
            self.test_textbox.insert("1.0", state["test_text"])
        if state["log"]:
            self._log(state["log"] + "\n")
        self._update_run_btn()

    def _section_header(self, parent, text):
        """セクションヘッダーを作成する。"""
        ctk.CTkLabel(
            parent, text=text,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("#4C566A", "#B0B8C8"),
        ).pack(anchor="w", padx=14, pady=(12, 6))

    # --- ファイル設定 ---
    def _build_file_section(self, parent):
        sec = ctk.CTkFrame(parent)
        sec.pack(fill="x", pady=(0, 12))

        self._section_header(sec, t("sec_file"))

        # 入力ファイル
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("input_file"), width=140, anchor="w").pack(side="left")
        self.input_var = ctk.StringVar()
        ctk.CTkEntry(row, textvariable=self.input_var).pack(side="left", fill="x", expand=True, padx=(4, 6))
        ctk.CTkButton(row, text=t("browse"), width=60, command=self._browse_input).pack(side="left")

        # 出力ファイル
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("output_file"), width=140, anchor="w").pack(side="left")
        self.output_var = ctk.StringVar()
        ctk.CTkEntry(row, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=(4, 6))
        ctk.CTkButton(row, text=t("browse"), width=60, command=self._browse_output).pack(side="left")

        # スライド範囲
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(3, 12))
        ctk.CTkLabel(row, text=t("slides"), width=100, anchor="w").pack(side="left")
        self.slide_range_var = ctk.StringVar(value="all")
        self._selected_slides: set[int] | None = None  # None = 全スライド
        ctk.CTkRadioButton(
            row, text=t("slides_all"), variable=self.slide_range_var, value="all",
            command=self._on_slide_range_changed,
        ).pack(side="left", padx=(0, 16))
        ctk.CTkRadioButton(
            row, text=t("slides_some"), variable=self.slide_range_var, value="select",
            command=self._on_slide_range_changed,
        ).pack(side="left", padx=(0, 8))
        self.slide_select_btn = ctk.CTkButton(
            row, text=t("slides_select"), width=60, command=self._open_slide_selector,
        )
        self.slide_select_label = ctk.CTkLabel(row, text="", anchor="w")
        # 初期状態では非表示
        self.slide_select_btn.pack_forget()
        self.slide_select_label.pack_forget()

    # --- 音声設定 ---
    def _build_voice_section(self, parent):
        sec = ctk.CTkFrame(parent)
        sec.pack(fill="x", pady=(0, 12))

        self._section_header(sec, t("sec_voice"))

        # エンジン選択
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("engine"), width=120, anchor="w").pack(side="left")
        self.engine_selector = ctk.CTkSegmentedButton(
            row, values=[self._engine_label(e) for e in self._ENGINES], command=self._on_engine_selected,
        )
        self.engine_selector.set(self._engine_label(self._engine))
        self.engine_selector.pack(side="left", padx=(4, 0))

        # URL
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        self.url_label = ctk.CTkLabel(row, text=t("voicevox_url"), width=120, anchor="w")
        self.url_label.pack(side="left")
        self.url_var = ctk.StringVar(value=self._engine_urls[self._engine])
        ctk.CTkEntry(row, textvariable=self.url_var).pack(side="left", fill="x", expand=True, padx=(4, 6))
        self.fetch_btn = ctk.CTkButton(row, text=t("fetch_speakers"), width=80, command=self._fetch_speakers)
        self.fetch_btn.pack(side="left")

        # モデル・APIキー (OpenAI互換のみ表示)
        self.openai_row = ctk.CTkFrame(sec, fg_color="transparent")
        ctk.CTkLabel(self.openai_row, text=t("model"), width=120, anchor="w").pack(side="left")
        self.model_var = ctk.StringVar(value="kokoro")
        ctk.CTkEntry(self.openai_row, textvariable=self.model_var, width=160).pack(side="left", padx=(4, 16))
        ctk.CTkLabel(self.openai_row, text=t("api_key"), anchor="w").pack(side="left")
        # APIキーは PPTX に残るため <config> には保存しない (環境変数 OPENAI_API_KEY を初期値にする)
        self.api_key_var = ctk.StringVar(value=os.environ.get("OPENAI_API_KEY", ""))
        ctk.CTkEntry(self.openai_row, textvariable=self.api_key_var, show="*",
                     placeholder_text=t("api_key_placeholder")).pack(side="left", fill="x", expand=True, padx=(4, 0))

        # 話者選択
        self.speaker_row = row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        self.speaker_label = ctk.CTkLabel(row, text=t("speaker"), width=120, anchor="w")
        self.speaker_label.pack(side="left")
        self.speaker_menu = ctk.CTkComboBox(
            row, values=[t("press_fetch")],
            command=self._on_speaker_changed, state="readonly",
        )
        self.speaker_menu.pack(side="left", fill="x", expand=True, padx=(4, 0))

        # スタイル選択 (VOICEVOX と、2段階に分けた OpenAI互換の声のときに表示)
        self.style_row = row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        self.style_label = ctk.CTkLabel(row, text=t("style"), width=120, anchor="w")
        self.style_label.pack(side="left")
        self.style_speaker_menu = ctk.CTkComboBox(
            row, values=["---"], state="readonly",
        )
        self.style_speaker_menu.pack(side="left", fill="x", expand=True, padx=(4, 0))

        # 読み上げ速度
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("speed"), width=120, anchor="w").pack(side="left")
        self.speed_var = ctk.DoubleVar(value=1.0)
        ctk.CTkSlider(row, from_=0.5, to=2.0, number_of_steps=30, variable=self.speed_var).pack(
            side="left", fill="x", expand=True, padx=(4, 8)
        )
        self.speed_label = ctk.CTkLabel(row, text="1.0", width=40)
        self.speed_label.pack(side="left")
        self.speed_var.trace_add("write", lambda *_: self.speed_label.configure(text=f"{self.speed_var.get():.1f}"))

        # ピッチ (VOICEVOXのみ表示)
        self.pitch_row = row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("pitch"), width=120, anchor="w").pack(side="left")
        self.pitch_var = ctk.DoubleVar(value=0.0)
        ctk.CTkSlider(row, from_=-0.15, to=0.15, number_of_steps=30, variable=self.pitch_var).pack(
            side="left", fill="x", expand=True, padx=(4, 8)
        )
        self.pitch_label = ctk.CTkLabel(row, text="0.00", width=40)
        self.pitch_label.pack(side="left")
        self.pitch_var.trace_add("write", lambda *_: self.pitch_label.configure(text=f"{self.pitch_var.get():.2f}"))

        # 抑揚 (VOICEVOXのみ表示)
        self.intonation_row = row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("intonation"), width=120, anchor="w").pack(side="left")
        self.intonation_var = ctk.DoubleVar(value=1.0)
        ctk.CTkSlider(row, from_=0.0, to=2.0, number_of_steps=20, variable=self.intonation_var).pack(
            side="left", fill="x", expand=True, padx=(4, 8)
        )
        self.intonation_label = ctk.CTkLabel(row, text="1.0", width=40)
        self.intonation_label.pack(side="left")
        self.intonation_var.trace_add("write", lambda *_: self.intonation_label.configure(text=f"{self.intonation_var.get():.1f}"))

        # 音量
        self.volume_row = row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("volume"), width=120, anchor="w").pack(side="left")
        self.volume_var = ctk.DoubleVar(value=1.0)
        ctk.CTkSlider(row, from_=0.0, to=2.0, number_of_steps=20, variable=self.volume_var).pack(
            side="left", fill="x", expand=True, padx=(4, 8)
        )
        self.volume_label = ctk.CTkLabel(row, text="1.0", width=40)
        self.volume_label.pack(side="left")
        self.volume_var.trace_add("write", lambda *_: self.volume_label.configure(text=f"{self.volume_var.get():.1f}"))

        # テスト再生
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("test"), width=120, anchor="nw").pack(side="left", anchor="n")
        self.test_textbox = ctk.CTkTextbox(row, height=60, wrap="word")
        self.test_textbox.insert("1.0", t("test_text"))
        self.test_textbox.pack(side="left", fill="x", expand=True, padx=(4, 6))
        self.test_textbox._textbox.configure(height=3)
        self.test_play_btn = ctk.CTkButton(
            row, text=t("play"), width=80, command=self._on_test_play,
            state="disabled",
        )
        self.test_play_btn.pack(side="left", anchor="n")

        # VOICEVOX 利用規約リンク (VOICEVOXのみ表示)
        self.voicevox_note_row = row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(0, 3))
        # 120px のスペーサーでラベル列と揃える
        ctk.CTkLabel(row, text="", width=120).pack(side="left")
        note = ctk.CTkLabel(
            row,
            text=t("voicevox_terms"),
            text_color=None,
            font=ctk.CTkFont(size=11),
            cursor="hand2",
        )
        note.pack(side="left", padx=(4, 0))
        note.bind("<Button-1>", lambda e: webbrowser.open("https://voicevox.hiroshiba.jp/"))

        # 文の区切り・末尾の余白
        self.pause_row = row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(3, 12))
        ctk.CTkLabel(row, text=t("pause"), width=120, anchor="w").pack(side="left")
        self.pause_var = ctk.DoubleVar(value=0.5)
        ctk.CTkEntry(row, textvariable=self.pause_var, width=50).pack(side="left", padx=(4, 16))
        ctk.CTkLabel(row, text=t("end_pause"), width=120, anchor="w").pack(side="left")
        self.end_pause_var = ctk.DoubleVar(value=2.0)
        ctk.CTkEntry(row, textvariable=self.end_pause_var, width=50).pack(side="left", padx=(4, 16))
        # 行の途中の文末 (. ! ?) でも区切る (英語のように改行なしの段落で書かれたノート向け)
        # 初期値はエンジンに合わせる (VOICEVOX: オフ, OpenAI互換: オン)
        self.split_sentence_var = ctk.BooleanVar(value=self._engine != "voicevox")
        ctk.CTkCheckBox(row, text=t("split_sentences"), variable=self.split_sentence_var).pack(side="left")


    # --- 字幕設定 ---
    def _build_subtitle_section(self, parent):
        sec = ctk.CTkFrame(parent)
        sec.pack(fill="x", pady=(0, 12))

        self._section_header(sec, t("sec_subtitle"))

        # 字幕有効
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        self.subtitle_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(row, text=t("show_subtitles"), variable=self.subtitle_var).pack(side="left")

        # 句読点の置換
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("punct_replace"), width=120, anchor="w").pack(side="left")
        ctk.CTkLabel(row, text=t("touten"), anchor="w").pack(side="left")
        # 変数には表示用の文言が入る (内部の値への変換は _punct_value / _punct_label)
        self.touten_mode_var = ctk.StringVar(value=self._punct_label(_TOUTEN_CHOICES, "そのまま"))
        ctk.CTkComboBox(row, values=[self._punct_label(_TOUTEN_CHOICES, v) for v, _ in _TOUTEN_CHOICES],
                        variable=self.touten_mode_var,
                        state="readonly", width=150).pack(side="left", padx=(4, 16))
        ctk.CTkLabel(row, text=t("kuten"), anchor="w").pack(side="left")
        self.kuten_mode_var = ctk.StringVar(value=self._punct_label(_KUTEN_CHOICES, "そのまま"))
        ctk.CTkComboBox(row, values=[self._punct_label(_KUTEN_CHOICES, v) for v, _ in _KUTEN_CHOICES],
                        variable=self.kuten_mode_var,
                        state="readonly", width=150).pack(side="left", padx=(4, 0))

        # 文字装飾 (太字・斜体・下線)
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("decoration"), width=120, anchor="w").pack(side="left")
        self.default_bold_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(row, text=t("bold"), variable=self.default_bold_var, width=60).pack(side="left", padx=(0, 12))
        self.default_italic_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(row, text=t("italic"), variable=self.default_italic_var, width=60).pack(side="left", padx=(0, 12))
        self.default_underline_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(row, text=t("underline"), variable=self.default_underline_var, width=60).pack(side="left", padx=(0, 12))
        self.math_bold_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(row, text=t("math_bold"), variable=self.math_bold_var, width=60).pack(side="left")

        # スタイル
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("sub_style"), width=120, anchor="w").pack(side="left")
        self.style_var = ctk.StringVar(value="outline")
        ctk.CTkRadioButton(
            row, text=t("style_outline"), variable=self.style_var, value="outline",
            command=self._on_style_changed,
        ).pack(side="left", padx=(0, 16))
        ctk.CTkRadioButton(
            row, text=t("style_box"), variable=self.style_var, value="box",
            command=self._on_style_changed,
        ).pack(side="left")

        # フォントサイズ
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("font_size"), width=120, anchor="w").pack(side="left")
        self.fontsize_var = ctk.IntVar(value=18)
        ctk.CTkSlider(row, from_=10, to=48, number_of_steps=38, variable=self.fontsize_var).pack(
            side="left", fill="x", expand=True, padx=(4, 8)
        )
        self.fontsize_label = ctk.CTkLabel(row, text="18", width=40)
        self.fontsize_label.pack(side="left")
        self.fontsize_var.trace_add(
            "write", lambda *_: self.fontsize_label.configure(text=str(self.fontsize_var.get()))
        )

        # フォント名
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("font"), width=120, anchor="w").pack(side="left")
        self._font_default_label = t("font_default")
        self.subtitle_font_var = ctk.StringVar(value=self._font_default_label)
        ctk.CTkEntry(row, textvariable=self.subtitle_font_var).pack(side="left", fill="x", expand=True, padx=(4, 6))
        ctk.CTkButton(row, text=t("select"), width=60, command=self._open_font_picker).pack(side="left")

        # 下マージン
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("bottom_margin"), width=120, anchor="w").pack(side="left")
        self.bottom_var = ctk.DoubleVar(value=0.05)
        ctk.CTkSlider(row, from_=0.0, to=0.3, number_of_steps=30, variable=self.bottom_var).pack(
            side="left", fill="x", expand=True, padx=(4, 8)
        )
        self.bottom_label = ctk.CTkLabel(row, text="0.05", width=40)
        self.bottom_label.pack(side="left")
        self.bottom_var.trace_add(
            "write", lambda *_: self.bottom_label.configure(text=f"{self.bottom_var.get():.2f}")
        )

        # 文字色 (共通)
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=t("font_color"), width=120, anchor="w").pack(side="left")
        self.font_color_var = ctk.StringVar(value="#FFFFFF")
        self.font_color_btn = ctk.CTkButton(
            row, text="#FFFFFF", width=90, fg_color="#FFFFFF", text_color="#000000",
            command=lambda: self._pick_color(self.font_color_var, self.font_color_btn),
        )
        self.font_color_btn.pack(side="left", padx=(4, 0))

        # --- 縁取りオプション (outline 時のみ表示) ---
        # 輪郭 (チェックボックス + 色 + 太さ)
        self.outline_row = ctk.CTkFrame(sec, fg_color="transparent")
        self.use_outline_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            self.outline_row, text=t("outline"), variable=self.use_outline_var, width=120,
        ).pack(side="left")
        self.outline_color_var = ctk.StringVar(value="#000000")
        self.outline_color_btn = ctk.CTkButton(
            self.outline_row, text="#000000", width=90, fg_color="#000000", text_color="#FFFFFF",
            command=lambda: self._pick_color(self.outline_color_var, self.outline_color_btn),
        )
        self.outline_color_btn.pack(side="left", padx=(4, 12))
        ctk.CTkLabel(self.outline_row, text=t("outline_width"), width=30, anchor="w").pack(side="left")
        self.outline_width_var = ctk.DoubleVar(value=0.75)
        ctk.CTkSlider(
            self.outline_row, from_=0.25, to=6.0, number_of_steps=23,
            variable=self.outline_width_var, width=120,
        ).pack(side="left", padx=(4, 8))
        self.outline_width_label = ctk.CTkLabel(self.outline_row, text="0.75", width=36)
        self.outline_width_label.pack(side="left")
        self.outline_width_var.trace_add(
            "write", lambda *_: self.outline_width_label.configure(text=f"{self.outline_width_var.get():.2f}")
        )

        # ぼかし (チェックボックス + 色 + サイズ)
        self.glow_row = ctk.CTkFrame(sec, fg_color="transparent")
        self.use_glow_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            self.glow_row, text=t("glow"), variable=self.use_glow_var, width=120,
        ).pack(side="left")
        self.glow_color_var = ctk.StringVar(value="#000000")
        self.glow_color_btn = ctk.CTkButton(
            self.glow_row, text="#000000", width=90, fg_color="#000000", text_color="#FFFFFF",
            command=lambda: self._pick_color(self.glow_color_var, self.glow_color_btn),
        )
        self.glow_color_btn.pack(side="left", padx=(4, 12))
        ctk.CTkLabel(self.glow_row, text=t("glow_size"), width=40, anchor="w").pack(side="left")
        self.glow_size_var = ctk.DoubleVar(value=11.0)
        ctk.CTkSlider(
            self.glow_row, from_=1.0, to=30.0, number_of_steps=29,
            variable=self.glow_size_var, width=120,
        ).pack(side="left", padx=(4, 8))
        self.glow_size_label = ctk.CTkLabel(self.glow_row, text="11.0", width=30)
        self.glow_size_label.pack(side="left")
        self.glow_size_var.trace_add(
            "write", lambda *_: self.glow_size_label.configure(text=f"{self.glow_size_var.get():.1f}")
        )

        # --- 背景オプション (box 時のみ表示) ---
        self.bg_row = ctk.CTkFrame(sec, fg_color="transparent")
        ctk.CTkLabel(self.bg_row, text=t("bg_color"), width=120, anchor="w").pack(side="left")
        self.bg_color_var = ctk.StringVar(value="#000000")
        self.bg_color_btn = ctk.CTkButton(
            self.bg_row, text="#000000", width=90, fg_color="#000000", text_color="#FFFFFF",
            command=lambda: self._pick_color(self.bg_color_var, self.bg_color_btn),
        )
        self.bg_color_btn.pack(side="left", padx=(4, 16))

        ctk.CTkLabel(self.bg_row, text=t("bg_alpha"), width=60, anchor="w").pack(side="left")
        self.bg_alpha_var = ctk.IntVar(value=60)
        ctk.CTkSlider(self.bg_row, from_=0, to=100, number_of_steps=100, variable=self.bg_alpha_var, width=160).pack(
            side="left", padx=(4, 8)
        )
        self.bg_alpha_label = ctk.CTkLabel(self.bg_row, text="60%", width=40)
        self.bg_alpha_label.pack(side="left")
        self.bg_alpha_var.trace_add(
            "write", lambda *_: self.bg_alpha_label.configure(text=f"{self.bg_alpha_var.get()}%")
        )

        # 初期表示状態を設定
        self._on_style_changed()

    # --- アニメーション ---
    def _build_animation_section(self, parent):
        sec = ctk.CTkFrame(parent)
        sec.pack(fill="x", pady=(0, 12))

        self._section_header(sec, t("sec_animation"))

        # <next> 余りアニメーション
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(3, 12))
        self.auto_next_enabled_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(row, text=t("auto_next"), variable=self.auto_next_enabled_var,
                         width=180).pack(side="left")
        self.auto_next_var = ctk.DoubleVar(value=5.0)
        ctk.CTkEntry(row, textvariable=self.auto_next_var, width=50).pack(side="left", padx=(4, 0))
        ctk.CTkLabel(row, text=t("auto_next_interval"), font=ctk.CTkFont(size=11), text_color="gray50").pack(side="left", padx=(4, 0))
        ctk.CTkLabel(row, text=t("auto_next_off"), font=ctk.CTkFont(size=11), text_color="gray50").pack(side="left", padx=(8, 0))
        ctk.CTkButton(row, text=t("check_next"), width=80, height=28,
                       command=self._check_next_tags).pack(side="right")

    # --- 実行 / ログ ---
    def _build_action_section(self, parent):
        sec = ctk.CTkFrame(parent)
        sec.pack(fill="both", expand=True)

        # ドキュメント・GitHub へのリンクと表示言語
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(14, 0))
        for text, url in ((t("link_docs"), DOCS_URL_EN if i18n.get_language() == "en" else DOCS_URL),
                          ("GitHub", GITHUB_URL)):
            link = ctk.CTkLabel(row, text=text, cursor="hand2", text_color=("#4F46E5", "#A5B4FC"),
                                font=ctk.CTkFont(underline=True))
            link.pack(side="left", padx=(0, 14))
            link.bind("<Button-1>", lambda e, u=url: webbrowser.open(u))
        self.language_selector = ctk.CTkSegmentedButton(
            row, values=list(i18n.LANGUAGES.values()), command=self._on_language_selected,
        )
        self.language_selector.set(i18n.LANGUAGES[i18n.get_language()])
        self.language_selector.pack(side="right")
        ctk.CTkLabel(row, text="言語 / Language").pack(side="right", padx=(0, 8))

        self.run_btn = ctk.CTkButton(
            sec, text=t("run"), font=ctk.CTkFont(size=15, weight="bold"),
            height=44, corner_radius=12, command=self._on_run,
        )
        self.run_btn.pack(fill="x", padx=14, pady=(14, 8))

        self.progress = ctk.CTkProgressBar(sec, height=6, corner_radius=3)
        self.progress.set(0)
        # 初期状態では非表示 (生成開始時に表示)

        self.log_box = ctk.CTkTextbox(sec, state="disabled", font=ctk.CTkFont(size=12))
        self.log_box.pack(fill="both", expand=True, padx=14, pady=(0, 14))

    # ------------------------------------------------------------------
    # コールバック
    # ------------------------------------------------------------------

    def _browse_input(self):
        path = filedialog.askopenfilename(filetypes=[("PowerPoint", "*.pptx")])
        if path:
            self._set_input_file(path)

    def _set_input_file(self, path: str):
        """入力ファイルを設定し、<config> タグを自動読み込みする。"""
        self.input_var.set(path)
        # 入力ファイルを指定するたびに出力ファイル名も追従させる
        base = os.path.splitext(path)[0]
        self.output_var.set(base + "_speech.pptx")
        # <config> / <cast> タグの自動読み込み (役はファイルごとに読み直す)
        try:
            config = pipeline.read_config(path)
            casts = pipeline.read_cast(path)
            self._voice_extra.cast = {}
            if config or casts:
                self._apply_config(config, casts)
                details = "\n".join([f"  {k}={v}" for k, v in config.items()]
                                    + [f"  <cast {name}>" for name in casts])
                self._log(t("log_config_loaded", details=details))
        except Exception:
            pass

    def _on_slide_range_changed(self):
        if self.slide_range_var.get() == "select":
            self.slide_select_btn.pack(side="left", padx=(0, 4))
            self.slide_select_label.pack(side="left")
        else:
            self.slide_select_btn.pack_forget()
            self.slide_select_label.pack_forget()
            self._selected_slides = None

    def _open_slide_selector(self):
        input_path = self.input_var.get().strip()
        if not input_path or not os.path.exists(input_path):
            self._log(t("log_need_input_for_select"))
            return
        try:
            slides = read_slides(input_path)
        except Exception as e:
            self._log(t("log_read_failed", e=e))
            return
        total = len(slides)
        if total == 0:
            self._log(t("log_no_slides"))
            return

        # ポップアップウィンドウ
        dialog = ctk.CTkToplevel(self)
        dialog.title(t("slide_select_title"))
        dialog.geometry("360x420")
        dialog.resizable(False, True)
        dialog.grab_set()

        ctk.CTkLabel(
            dialog, text=t("slide_select_heading", total=total),
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(padx=10, pady=(10, 4))

        # 全選択/全解除ボタン
        btn_row = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_row.pack(fill="x", padx=10, pady=(0, 4))

        check_vars: list[ctk.BooleanVar] = []

        def select_all():
            for v in check_vars:
                v.set(True)

        def deselect_all():
            for v in check_vars:
                v.set(False)

        ctk.CTkButton(btn_row, text=t("select_all"), width=70, command=select_all).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_row, text=t("deselect_all"), width=70, command=deselect_all).pack(side="left")

        # チェックボックスリスト
        scroll = ctk.CTkScrollableFrame(dialog)
        scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        prev = self._selected_slides
        for i in range(total):
            num = i + 1
            var = ctk.BooleanVar(value=(prev is None or num in prev))
            check_vars.append(var)
            notes = slides[i].notes_text or ""
            preview = notes.replace("\n", " ")[:30]
            label = t("slide_n", num=num)
            if preview:
                label += f" - {preview}"
            ctk.CTkCheckBox(scroll, text=label, variable=var).pack(anchor="w", pady=1)

        # OKボタン
        def on_ok():
            selected = {i + 1 for i, v in enumerate(check_vars) if v.get()}
            if not selected:
                messagebox.showwarning(t("no_selection_title"), t("no_selection"), parent=dialog)
                return
            self._selected_slides = selected
            self._update_slide_label(total)
            dialog.destroy()

        ctk.CTkButton(
            dialog, text="OK", width=120, command=on_ok,
        ).pack(pady=(0, 10))

    def _update_slide_label(self, total: int):
        """選択されたスライド番号をラベルに表示する。"""
        sel = self._selected_slides
        if sel is None or len(sel) == total:
            self.slide_select_label.configure(text="")
            return
        nums = sorted(sel)
        # 連番をまとめる: [1,2,3,5,7,8] → "1-3, 5, 7-8"
        parts = []
        start = nums[0]
        end = nums[0]
        for n in nums[1:]:
            if n == end + 1:
                end = n
            else:
                parts.append(f"{start}-{end}" if start != end else str(start))
                start = end = n
        parts.append(f"{start}-{end}" if start != end else str(start))
        self.slide_select_label.configure(text=", ".join(parts))

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            filetypes=[("PowerPoint", "*.pptx")], defaultextension=".pptx",
        )
        if path:
            self.output_var.set(path)

    # ------------------------------------------------------------------
    # エンジン切り替え
    # ------------------------------------------------------------------

    @staticmethod
    def _engine_label(engine: str) -> str:
        return "VOICEVOX" if engine == "voicevox" else t("engine_openai")

    def _on_engine_selected(self, label: str):
        engine = next(e for e in self._ENGINES if self._engine_label(e) == label)
        self._set_engine(engine)

    def _set_engine(self, engine: str):
        """TTS エンジンを切り替え、エンジンに応じて表示する項目を変える。"""
        if engine not in self._ENGINES:
            return
        self.engine_selector.set(self._engine_label(engine))
        if engine == self._engine:
            return
        # URL はエンジンごとに覚えておく
        self._engine_urls[self._engine] = self.url_var.get().strip()
        self._engine = engine
        self.url_var.set(self._engine_urls[engine])
        self._wanted_voice = ("", "")
        self._openai_voices = []
        self._voices_grouped = False
        self._apply_engine_ui()

        # 文末で区切るかはエンジンの想定言語に合わせる (<config> の split_sentences で上書き可)
        self.split_sentence_var.set(engine != "voicevox")

        # 話者一覧はエンジンごとに取り直す
        self._speakers_cache = []
        self._speaker_map.clear()
        self._styles_by_speaker = {}
        self.speaker_menu.configure(values=[t("press_fetch")])
        self.speaker_menu.set(t("press_fetch"))
        self.style_speaker_menu.configure(values=["---"])
        self.style_speaker_menu.set("---")
        self._update_run_btn()

    def _apply_engine_ui(self):
        """現在のエンジンに応じて、表示する項目とラベルを切り替える。"""
        is_voicevox = self._engine == "voicevox"
        self.engine_selector.set(self._engine_label(self._engine))
        self.url_label.configure(text=t("voicevox_url") if is_voicevox else t("url"))
        grouped = not is_voicevox and self._voices_grouped
        self.speaker_label.configure(
            text=t("speaker") if is_voicevox else (t("voice_group") if grouped else t("voice")))
        self.style_label.configure(text=t("style") if is_voicevox else t("voice"))
        self.fetch_btn.configure(text=t("fetch_speakers") if is_voicevox else t("fetch_voices"))
        if is_voicevox:
            self.openai_row.pack_forget()
            self.style_row.pack(fill="x", padx=14, pady=3, after=self.speaker_row)
            self.voicevox_note_row.pack(fill="x", padx=14, pady=(0, 3), before=self.pause_row)
            # ピッチ・抑揚は VOICEVOX 専用
            self.pitch_row.pack(fill="x", padx=14, pady=3, before=self.volume_row)
            self.intonation_row.pack(fill="x", padx=14, pady=3, before=self.volume_row)
        else:
            self.openai_row.pack(fill="x", padx=14, pady=3, before=self.speaker_row)
            if grouped:
                self.style_row.pack(fill="x", padx=14, pady=3, after=self.speaker_row)
            else:
                self.style_row.pack_forget()
            self.voicevox_note_row.pack_forget()
            self.pitch_row.pack_forget()
            self.intonation_row.pack_forget()

    def _create_engine(self, settings: Settings):
        """設定から TTS エンジンと役の声を作る (取得済みの話者一覧を使う)。"""
        speakers = self._speakers_cache if settings.engine == "voicevox" else self._openai_voices
        return pipeline.create_engine(settings), pipeline.create_cast(settings, speakers=speakers)

    def _current_voice_id(self) -> str:
        """OpenAI互換で選択中の声の ID。"""
        if self._voices_grouped:
            return str(self._speaker_map.get(self.style_speaker_menu.get(), ""))
        return self.speaker_menu.get()

    @staticmethod
    def _voice_group_label(key: tuple[str, str] | None) -> str:
        """Kokoro の声のグループ (言語の記号, 性別の記号) の表示名。"""
        if key is None:
            return t("voice_group_other")
        lang, gender = key
        lang_name = t(f"kokoro_lang_{lang}") if lang in KOKORO_LANG_CODES else t("lang_unknown", code=lang)
        return t("voice_group_label", lang=lang_name, gender=t(f"gender_{gender}"))

    def _current_voice_names(self) -> tuple[str, str]:
        """選択中の話者の (話者名, スタイル名)。OpenAI互換は (声の ID, "")。一覧の取得前は選びたい話者。"""
        if not self._speaker_map:
            return self._wanted_voice
        if self._engine == "openai":
            return self._current_voice_id(), ""
        speaker = self.speaker_menu.get()
        style = self.style_speaker_menu.get()
        return (speaker.rsplit(" (", 1)[0] if " (" in speaker else "",
                style.rsplit(" (ID=", 1)[0] if " (ID=" in style else "")

    # 話者取得のタイムアウト (接続, 読み込み) 秒
    _FETCH_TIMEOUT = pipeline.FETCH_TIMEOUT

    def _fetch_speakers(self):
        """話者一覧を別スレッドで取得する (取得中も GUI が固まらないように)。"""
        self._log_clear()
        # 取得し直しても、今選んでいる話者を選び直す
        if self._speaker_map:
            self._wanted_voice = self._current_voice_names()
        url = self.url_var.get().strip().rstrip("/")
        name = t("engine_name_voicevox") if self._engine == "voicevox" else t("engine_name_server")
        self._log(t("log_connecting", name=name, url=url))
        self.fetch_btn.configure(text=t("fetching"), state="disabled")
        engine = self._engine
        api_key = self.api_key_var.get().strip()
        threading.Thread(target=self._fetch_speakers_worker, args=(url, engine, api_key), daemon=True).start()

    def _fetch_speakers_worker(self, url: str, engine: str = "voicevox", api_key: str = ""):
        settings = Settings()
        settings.set_engine(engine)
        settings.url = url
        settings.api_key = api_key
        try:
            speakers = pipeline.fetch_speakers(settings)
            if engine == "openai":
                self.after(0, self._on_voices_fetched, engine, speakers)
                return
        except requests.exceptions.ConnectTimeout:
            msg = t("log_connect_timeout", url=url, sec=self._FETCH_TIMEOUT[0])
        except requests.exceptions.ConnectionError:
            msg = t("log_connect_failed", url=url)
        except requests.exceptions.Timeout:
            msg = t("log_read_timeout", url=url, sec=self._FETCH_TIMEOUT[1])
        except Exception as e:
            msg = t("log_fetch_failed", e=e)
        else:
            self.after(0, self._on_speakers_fetched, speakers)
            return
        name = t("engine_name_voicevox") if engine == "voicevox" else t("engine_name_server")
        self.after(0, self._on_speakers_fetch_failed, msg + t("log_check_server", name=name))

    def _reset_fetch_btn(self):
        self.fetch_btn.configure(text=t("fetch_speakers") if self._engine == "voicevox" else t("fetch_voices"),
                                 state="normal")

    def _on_speakers_fetch_failed(self, msg: str):
        self._reset_fetch_btn()
        self._log(msg)

    def _on_voices_fetched(self, engine: str, voices: list[str], quiet: bool = False):
        """OpenAI互換サーバの声一覧を反映する。quiet なら結果をログに出さない (画面の作り直し時)。"""
        self._reset_fetch_btn()
        if engine != self._engine:
            return  # 取得中にエンジンが切り替えられた
        self._openai_voices = list(voices)
        groups = group_voices(voices)
        self._voices_grouped = groups is not None
        self._styles_by_speaker = {}
        if groups:
            # Kokoro: 1段目 = 言語・性別 (VOICEVOX の話者の欄), 2段目 = 声 (スタイルの欄)
            for key, ids in groups:
                self._styles_by_speaker[self._voice_group_label(key)] = [
                    (f"{voice_short_name(v)} ({v})", v) for v in ids]
            names = [f"{label} ({len(v)})" for label, v in self._styles_by_speaker.items()]
            self.speaker_menu.configure(values=names)
            self.speaker_menu.set(names[0])
            self._on_speaker_changed(names[0])
        else:
            self._speaker_map = {v: v for v in voices}
            self.style_speaker_menu.configure(values=["---"])
            self.style_speaker_menu.set("---")
            if voices:
                self.speaker_menu.configure(values=voices)
                self.speaker_menu.set(voices[0])
        self._apply_engine_ui()
        if voices:
            if not quiet:
                self._log(t("log_voices_fetched", n=len(voices)))
            self._select_voice(*self._wanted_voice)
        else:
            self._log(t("log_no_voices"))
        self._update_run_btn()

    def _on_speakers_fetched(self, speakers: list[dict], quiet: bool = False):
        """VOICEVOX の話者一覧を反映する。quiet なら結果をログに出さない (画面の作り直し時)。"""
        self._reset_fetch_btn()
        if self._engine != "voicevox":
            return  # 取得中にエンジンが切り替えられた
        self._speakers_cache = speakers
        self._speaker_map.clear()
        # 話者名 → [(スタイルラベル, ID), ...] のマッピング
        self._styles_by_speaker: dict[str, list[tuple[str, int]]] = {}
        speaker_names = []
        total_styles = 0
        for sp in speakers:
            name = sp["name"]
            styles = []
            for style in sp.get("styles", []):
                label = f"{style['name']} (ID={style['id']})"
                styles.append((label, style["id"]))
                total_styles += 1
            if styles:
                self._styles_by_speaker[name] = styles
                speaker_names.append(f"{name} ({len(styles)})")

        if speaker_names:
            self.speaker_menu.configure(values=speaker_names)
            self.speaker_menu.set(speaker_names[0])
            self._on_speaker_changed(speaker_names[0])
            if not quiet:
                self._log(t("log_speakers_fetched", n=len(speakers), styles=total_styles))
            self._select_voice(*self._wanted_voice)
        else:
            self._log(t("log_no_speakers"))
        self._update_run_btn()

    def _on_speaker_changed(self, speaker_name: str):
        if self._engine != "voicevox" and not self._voices_grouped:
            return  # 2段階に分けていない OpenAI互換の声にはスタイルの欄がない
        # "話者名 (N)" → "話者名" に変換
        name = speaker_name.rsplit(" (", 1)[0] if " (" in speaker_name else speaker_name
        styles = self._styles_by_speaker.get(name, [])
        self._speaker_map.clear()
        labels = []
        for label, sid in styles:
            labels.append(label)
            self._speaker_map[label] = sid
        if labels:
            self.style_speaker_menu.configure(values=labels)
            self.style_speaker_menu.set(labels[0])
        else:
            self.style_speaker_menu.configure(values=["---"])
            self.style_speaker_menu.set("---")

    # ------------------------------------------------------------------
    # <next> / アニメーション チェック
    # ------------------------------------------------------------------

    def _check_next_tags(self):
        """各スライドのクリックアニメーション数と <next> タグ数を比較してログに出力する。"""
        path = self.input_var.get()
        if not path or not os.path.isfile(path):
            self._log(t("log_no_input"))
            return
        try:
            rows = pipeline.check_next_tags(path)
        except Exception as e:
            self._log(t("log_pptx_read_failed", e=e))
            return

        self._log(t("log_next_check_header"))
        for r in rows:
            self._log(t("log_next_check_row", num=r.slide, anims=r.animations, nexts=r.nexts, status=r.status))
        self._log("---\n")

    # ------------------------------------------------------------------
    # テスト再生
    # ------------------------------------------------------------------

    def _on_test_play(self):
        btn_text = self.test_play_btn.cget("text")
        if btn_text == t("stop_play"):
            self._test_stop = True
            winsound.PlaySound(None, winsound.SND_PURGE)
            self._test_play_reset()
            return

        text = self.test_textbox.get("1.0", "end-1c").strip()
        if not text:
            return
        if not self._speaker_map:
            self._log(t("log_fetch_first"))
            return

        settings = self._settings_from_ui()
        engine, cast = self._create_engine(settings)

        self.test_play_btn.configure(text=t("synthesizing"), state="disabled")
        thread = threading.Thread(
            target=self._test_play_worker,
            args=(text, settings, engine, cast),
            daemon=True,
        )
        thread.start()

    def _test_play_worker(self, text, settings, engine, cast):
        try:
            wav, timings, _ = pipeline.synthesize_notes(engine, text, cast, settings.pause,
                                                        settings.split_sentences, turn_pause=settings.turn_pause)
            self.after(0, lambda: self.test_play_btn.configure(text=t("stop_play"), state="normal"))

            # 字幕を再生タイミングに合わせてログに表示
            self._test_stop = False
            if timings:
                self.after(0, lambda: self._log(t("log_subtitle_preview")))
                sub_thread = threading.Thread(
                    target=self._test_subtitle_worker,
                    args=(timings,), daemon=True,
                )
                sub_thread.start()

            # SND_MEMORY は同期再生 (再生完了 or SND_PURGE で停止するまでブロック)
            winsound.PlaySound(wav, winsound.SND_MEMORY)
            self._test_stop = True
        except Exception as e:
            self._test_stop = True
            # except を抜けると e は消えるため、値を先に束縛する
            self.after(0, lambda msg=str(e): self._log(t("log_test_error", msg=msg)))
        finally:
            self.after(0, self._test_play_reset)

    _STRIP_TAGS = re.compile(r"</?[a-zA-Z][^>]*>")

    def _test_subtitle_worker(self, timings):
        """再生タイミングに合わせて字幕テキストをログに表示する。"""
        import time
        t0 = time.perf_counter()
        for timing in timings:
            if self._test_stop:
                return
            # 開始タイミングまで待つ
            wait = timing.start_ms / 1000.0 - (time.perf_counter() - t0)
            if wait > 0:
                time.sleep(wait)
            if self._test_stop:
                return
            clean = self._STRIP_TAGS.sub("", timing.text)
            clean = _unescape_display(clean)
            self.after(0, lambda t=clean: self._log(f"  {t}\n"))

    def _test_play_reset(self):
        state = "normal" if self._speaker_map else "disabled"
        self.test_play_btn.configure(text=t("play"), state=state)

    # ------------------------------------------------------------------
    # <config> タグ
    # ------------------------------------------------------------------

    def _settings_from_ui(self) -> Settings:
        """画面の値から Settings を作る。"""
        s = Settings()
        s.engine = self._engine
        s.urls = dict(self._engine_urls)
        s.urls[self._engine] = self.url_var.get().strip()
        s.model = self.model_var.get().strip()
        s.api_key = self.api_key_var.get().strip()
        s.speaker, s.style = self._current_voice_names()
        s.speed = self.speed_var.get()
        s.pitch = self.pitch_var.get()
        s.intonation = self.intonation_var.get()
        s.volume = self.volume_var.get()
        s.pause = self.pause_var.get()
        s.end_pause = self.end_pause_var.get()
        s.split_sentences = self.split_sentence_var.get()
        s.auto_next = self.auto_next_var.get()
        s.auto_next_enabled = self.auto_next_enabled_var.get()
        s.subtitle = self.subtitle_var.get()
        s.subtitle_style = self.style_var.get()
        s.fontsize = self.fontsize_var.get()
        font = self.subtitle_font_var.get().strip()
        s.font = "" if font == self._font_default_label else font
        s.bottom = self.bottom_var.get()
        s.font_color = self.font_color_var.get()
        s.outline = self.use_outline_var.get()
        s.outline_color = self.outline_color_var.get()
        s.outline_width = self.outline_width_var.get()
        s.glow = self.use_glow_var.get()
        s.glow_color = self.glow_color_var.get()
        s.glow_size = self.glow_size_var.get()
        s.bg_color = self.bg_color_var.get()
        s.bg_alpha = self.bg_alpha_var.get()
        s.kuten = self._punct_value(_KUTEN_CHOICES, self.kuten_mode_var.get())
        s.touten = self._punct_value(_TOUTEN_CHOICES, self.touten_mode_var.get())
        s.bold = self.default_bold_var.get()
        s.italic = self.default_italic_var.get()
        s.underline = self.default_underline_var.get()
        s.math_bold = self.math_bold_var.get()
        s.role = self._voice_extra.role
        s.cast = dict(self._voice_extra.cast)
        s.turn_pause = self._voice_extra.turn_pause
        return s

    def _load_settings_to_ui(self, s: Settings):
        """Settings の値を画面に反映する。"""
        # エンジンは先に切り替える (切り替えると話者一覧がリセットされるため)
        self._set_engine(s.engine)
        self._engine_urls = dict(s.urls)
        self.url_var.set(s.url)
        self.model_var.set(s.model)
        self.api_key_var.set(s.api_key)
        self.speed_var.set(s.speed)
        self.pitch_var.set(s.pitch)
        self.intonation_var.set(s.intonation)
        self.volume_var.set(s.volume)
        self.pause_var.set(s.pause)
        self.end_pause_var.set(s.end_pause)
        self.split_sentence_var.set(s.split_sentences)
        self.auto_next_var.set(s.auto_next)
        self.auto_next_enabled_var.set(s.auto_next_enabled)
        self.subtitle_var.set(s.subtitle)
        self.style_var.set(s.subtitle_style)
        self._on_style_changed()
        self.fontsize_var.set(s.fontsize)
        self.subtitle_font_var.set(s.font if s.font else self._font_default_label)
        self.bottom_var.set(s.bottom)
        self._set_color(s.font_color, self.font_color_var, self.font_color_btn)
        self.use_outline_var.set(s.outline)
        self._set_color(s.outline_color, self.outline_color_var, self.outline_color_btn)
        self.outline_width_var.set(s.outline_width)
        self.use_glow_var.set(s.glow)
        self._set_color(s.glow_color, self.glow_color_var, self.glow_color_btn)
        self.glow_size_var.set(s.glow_size)
        self._set_color(s.bg_color, self.bg_color_var, self.bg_color_btn)
        self.bg_alpha_var.set(s.bg_alpha)
        self.kuten_mode_var.set(self._punct_label(_KUTEN_CHOICES, s.kuten))
        self.touten_mode_var.set(self._punct_label(_TOUTEN_CHOICES, s.touten))
        self.default_bold_var.set(s.bold)
        self.default_italic_var.set(s.italic)
        self.default_underline_var.set(s.underline)
        self.math_bold_var.set(s.math_bold)
        self._voice_extra.role = s.role
        self._voice_extra.cast = dict(s.cast)
        self._voice_extra.turn_pause = s.turn_pause
        # 話者は名前で選ぶ (一覧を取得していなければ、取得したときに選ぶ)
        self._wanted_voice = (s.speaker, s.style)
        if self._speaker_map:
            self._select_voice(s.speaker, s.style)

    def _apply_config(self, config: dict, casts: dict | None = None):
        """解析済みの <config> / <cast> を画面に反映する。"""
        s = self._settings_from_ui()
        s.apply_config(config, warn=lambda k, v: self._log(t("log_config_invalid", key=k, value=v)))
        for name, values in (casts or {}).items():
            s.apply_cast({name: values}, warn=lambda k, v, name=name: self._log(
                t("log_cast_invalid", name=name, key=k, value=v)))
        self._load_settings_to_ui(s)

    def _set_color(self, hex_val: str, var: ctk.StringVar, btn: ctk.CTkButton):
        """色の変数とボタン表示を更新する。"""
        if not hex_val.startswith("#"):
            hex_val = "#" + hex_val
        hex_val = hex_val.upper()
        var.set(hex_val)
        btn.configure(text=hex_val, fg_color=hex_val)
        try:
            r, g, b = int(hex_val[1:3], 16), int(hex_val[3:5], 16), int(hex_val[5:7], 16)
            text_col = "#000000" if (r * 0.299 + g * 0.587 + b * 0.114) > 128 else "#FFFFFF"
            btn.configure(text_color=text_col)
        except ValueError:
            pass

    def _select_voice(self, speaker: str, style: str):
        """話者一覧から、名前で話者 (とスタイル) を選ぶ。見つからなければ今の選択のまま。"""
        if self._engine == "openai":
            # OpenAI互換: speaker は声の ID (2段階の場合は、その声を含むグループを選ぶ)
            voice_id = speaker
            if not voice_id:
                return
            if not self._voices_grouped:
                if voice_id in self._speaker_map:
                    self.speaker_menu.set(voice_id)
                return
            for display_name in (self.speaker_menu.cget("values") or []):
                styles = self._styles_by_speaker.get(display_name.rsplit(" (", 1)[0], [])
                label = next((lb for lb, v in styles if v == voice_id), None)
                if label:
                    self.speaker_menu.set(display_name)
                    self._on_speaker_changed(display_name)
                    self.style_speaker_menu.set(label)
                    break
            return

        if speaker:
            for display_name in (self.speaker_menu.cget("values") or []):
                if display_name.rsplit(" (", 1)[0] == speaker:
                    self.speaker_menu.set(display_name)
                    self._on_speaker_changed(display_name)
                    break
        if style:
            for label in (self.style_speaker_menu.cget("values") or []):
                # "スタイル名 (ID=X)" から先頭のスタイル名を取得
                if label.rsplit(" (ID=", 1)[0] == style:
                    self.style_speaker_menu.set(label)
                    break

    def _generate_config_tag(self) -> str:
        """現在の GUI 設定を <config ...> タグ (と役の数だけの <cast ...> タグ) の文字列として生成する。"""
        s = self._settings_from_ui()
        return "\n".join([s.to_config_tag(), *s.to_cast_tags()])

    def _on_save_config(self):
        """設定保存ポップアップを表示する。"""
        tag = self._generate_config_tag()

        dialog = ctk.CTkToplevel(self)
        dialog.title(t("save_config"))
        dialog.geometry("600x220")
        dialog.resizable(True, False)
        dialog.grab_set()

        ctk.CTkLabel(
            dialog,
            text=t("save_config_desc"),
            font=ctk.CTkFont(size=12),
            justify="left",
        ).pack(padx=16, pady=(16, 8), anchor="w")

        text_box = ctk.CTkTextbox(dialog, height=80, font=ctk.CTkFont(size=11), wrap="word")
        text_box.pack(fill="x", padx=16, pady=(0, 8))
        text_box.insert("1.0", tag)
        text_box.tag_add("sel", "1.0", "end-1c")
        text_box.configure(state="disabled")

        def on_copy():
            self.clipboard_clear()
            self.clipboard_append(tag)
            copy_btn.configure(text=t("copied"))
            dialog.after(1500, lambda: copy_btn.configure(text=t("copy")))

        copy_btn = ctk.CTkButton(dialog, text=t("copy"), width=120, command=on_copy)
        copy_btn.pack(pady=(0, 16))

    def _open_font_picker(self):
        families = sorted(
            (f for f in set(tkfont.families()) if not f.startswith("@")),
            key=str.lower,
        )
        all_fonts = [self._font_default_label] + families

        dialog = ctk.CTkToplevel(self)
        dialog.title(t("font_picker_title"))
        dialog.geometry("400x500")
        dialog.resizable(True, True)
        dialog.grab_set()

        # 検索
        search_var = ctk.StringVar()
        ctk.CTkEntry(
            dialog, textvariable=search_var, placeholder_text=t("search"),
        ).pack(fill="x", padx=10, pady=(10, 6))

        # リスト
        list_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        listbox = tk.Listbox(list_frame, font=("", 11), activestyle="dotbox")
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True)

        def populate(query=""):
            listbox.delete(0, "end")
            q = query.lower()
            for f in all_fonts:
                if q in f.lower():
                    listbox.insert("end", f)
            # 現在の選択をハイライト
            current = self.subtitle_font_var.get()
            items = listbox.get(0, "end")
            if current in items:
                idx = list(items).index(current)
                listbox.selection_set(idx)
                listbox.see(idx)

        populate()
        search_var.trace_add("write", lambda *_: populate(search_var.get()))

        def on_ok():
            sel = listbox.curselection()
            if sel:
                self.subtitle_font_var.set(listbox.get(sel[0]))
            dialog.destroy()

        listbox.bind("<Double-1>", lambda e: on_ok())

        ctk.CTkButton(dialog, text="OK", width=120, command=on_ok).pack(pady=(0, 10))

    @staticmethod
    def _punct_label(choices, value: str) -> str:
        """句読点の置換先の内部の値 (または日英いずれかの表示名) → 現在の言語の表示名。"""
        for v, key in choices:
            names = (v,) if key is None else (v, *i18n.STRINGS[key])
            if value in names:
                return v if key is None else t(key)
        return value  # 一覧にない 1 文字の指定などはそのまま

    @staticmethod
    def _punct_value(choices, label: str) -> str:
        """表示名 → 句読点の置換先の内部の値。"""
        for v, key in choices:
            if label == v or (key is not None and label in i18n.STRINGS[key]):
                return v
        return label

    def _on_style_changed(self):
        if self.style_var.get() == "outline":
            self.bg_row.pack_forget()
            self.outline_row.pack(fill="x", padx=14, pady=3)
            self.glow_row.pack(fill="x", padx=14, pady=3)
        else:
            self.outline_row.pack_forget()
            self.glow_row.pack_forget()
            self.bg_row.pack(fill="x", padx=14, pady=3)

    def _pick_color(self, var: ctk.StringVar, btn: ctk.CTkButton):
        color = colorchooser.askcolor(color=var.get(), title=t("color_picker_title"))
        if color[1]:
            hex_color = color[1].upper()
            var.set(hex_color)
            btn.configure(text=hex_color, fg_color=hex_color)
            # テキストが見えるように明暗で文字色を切替
            r, g, b = color[0]
            text_col = "#000000" if (r * 0.299 + g * 0.587 + b * 0.114) > 128 else "#FFFFFF"
            btn.configure(text_color=text_col)

    def _log(self, text: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _log_clear(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ------------------------------------------------------------------
    # 生成処理
    # ------------------------------------------------------------------

    def _update_run_btn(self):
        """入力ファイル・話者選択の状態に応じてボタンの有効/無効を切り替える。"""
        if self._running:
            return
        input_ok = bool(self.input_var.get().strip())
        speaker_ok = bool(self._speaker_map)
        if input_ok and speaker_ok:
            self.run_btn.configure(state="normal")
        else:
            self.run_btn.configure(state="disabled")
        # テスト再生ボタン (再生中/合成中でなければ話者の有無で制御)
        btn_text = self.test_play_btn.cget("text")
        if btn_text == t("play"):
            self.test_play_btn.configure(state="normal" if speaker_ok else "disabled")

    def _on_run(self):
        if self._running:
            # 停止要求
            self._cancel_event.set()
            self.run_btn.configure(state="disabled", text=t("stopping"))
            return

        input_path = self.input_var.get().strip()
        if not input_path:
            self._log(t("log_specify_input"))
            return
        if not os.path.exists(input_path):
            self._log(t("log_file_not_found", path=input_path))
            return

        # 出力ファイルの上書き確認
        base_name = os.path.splitext(input_path)[0]
        output_path = self.output_var.get().strip()
        if not output_path:
            output_path = base_name + "_speech.pptx"
        existing = [output_path] if os.path.exists(output_path) else []
        if existing:
            names = "\n".join(os.path.basename(f) for f in existing)
            if not messagebox.askyesno(t("overwrite_title"), t("overwrite_msg", names=names)):
                return

        self._running = True
        self._cancel_event.clear()
        self._log_clear()
        self.run_btn.configure(text=t("stop"), fg_color="#EF4444", hover_color="#DC2626")
        self.progress.set(0)
        self.progress.pack(fill="x", padx=14, pady=(0, 8), before=self.log_box)

        thread = threading.Thread(target=self._run_generate, daemon=True)
        thread.start()

    def _run_generate(self):
        old_stdout = sys.stdout
        sys.stdout = LogRedirector(self.log_box)
        try:
            self._do_generate()
        except _CancelledError:
            print(t("log_cancelled"))
        except Exception as e:
            print(t("log_error", e=e))
        finally:
            sys.stdout = old_stdout
            self.after(0, self._on_done)

    def _on_done(self):
        self._running = False
        self._cancel_event.clear()
        self.progress.pack_forget()
        self.run_btn.configure(
            text=t("run"),
            fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"],
            hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"],
        )
        self._update_run_btn()

    def _do_generate(self):
        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip()
        if not output_path:
            output_path = os.path.splitext(input_path)[0] + "_speech.pptx"

        settings = self._settings_from_ui()
        engine, cast = self._create_engine(settings)
        if pipeline.generate(
            settings, input_path, output_path,
            selected_slides=self._selected_slides,
            engine=engine, cast=cast,
            on_progress=lambda v: self.after(0, self.progress.set, v),
            should_cancel=self._cancel_event.is_set,
        ):
            print(t("log_video_howto"))

if __name__ == "__main__":
    app = App()
    app.mainloop()
