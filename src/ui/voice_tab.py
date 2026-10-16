"""音声タブ: エンジン・キャスト (役の一覧と、選択中の役の設定)・間・辞書・テスト再生"""

import copy
import os
import re
import threading
import webbrowser
import winsound
from tkinter import messagebox

import customtkinter as ctk
import requests

import pipeline
from i18n import t
from settings import DEFAULT_URLS, ENGINES, Role, Settings
from tts.openai_compat import voice_short_name
from ui import widgets as w

# 役を追加するときに順に割り当てる字幕色 (見分けやすい色)
PALETTE = ["#A6E22E", "#FF7FBF", "#66D9EF", "#FD971F", "#AE81FF", "#E6DB74", "#F92672", "#78DCE8"]

_SPEAKER_TAGS = re.compile(r"<(?:speaker|style)=[^>]*>", re.IGNORECASE)
_BAD_NAME = re.compile(r'[<>"]')


def engine_label(engine: str) -> str:
    return "VOICEVOX" if engine == "voicevox" else t("engine_openai")


class VoiceTab(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.engine = "voicevox"
        self.urls = dict(DEFAULT_URLS)
        self.speakers_cache: list[dict] = []   # VOICEVOX の話者一覧
        self.voices: list[str] = []            # OpenAI互換の声の一覧
        self.roles: list[Role] = [Role("")]
        self.default_idx = 0
        self.selected = 0
        self.usage: dict[str | None, int] = {}  # ノートでの文の数 (None は <speaker> のない文)
        self.dictionary: dict[str, str] = {}
        self._loading = False
        self._build()

    # ------------------------------------------------------------------
    # 画面
    # ------------------------------------------------------------------

    def _build(self):
        # エンジンと URL
        r = w.row(self, pady=(8, 2))
        w.label(r, t("engine"), width=70)
        self.engine_selector = ctk.CTkSegmentedButton(
            r, values=[engine_label(e) for e in ENGINES], command=self._on_engine_selected)
        self.engine_selector.set(engine_label(self.engine))
        self.engine_selector.pack(side="left", padx=(4, 12))
        self.url_label = ctk.CTkLabel(r, text="URL")
        self.url_label.pack(side="left")
        self.url_var = ctk.StringVar(value=self.urls[self.engine])
        ctk.CTkEntry(r, textvariable=self.url_var).pack(side="left", fill="x", expand=True, padx=(4, 6))
        self.fetch_btn = ctk.CTkButton(r, text=t("fetch_speakers"), width=90, command=self.fetch)
        self.fetch_btn.pack(side="left")

        # モデル・API キー (OpenAI互換のみ)
        self.openai_row = r = ctk.CTkFrame(self, fg_color="transparent")
        w.label(r, t("model"), width=70)
        self.model_var = ctk.StringVar(value="kokoro")
        ctk.CTkEntry(r, textvariable=self.model_var, width=140).pack(side="left", padx=(4, 16))
        ctk.CTkLabel(r, text=t("api_key")).pack(side="left")
        # API キーは PPTX に残るため <config> には保存しない (環境変数 OPENAI_API_KEY を初期値にする)
        self.api_key_var = ctk.StringVar(value=os.environ.get("OPENAI_API_KEY", ""))
        ctk.CTkEntry(r, textvariable=self.api_key_var, show="*",
                     placeholder_text=t("api_key_placeholder")).pack(side="left", fill="x", expand=True, padx=(4, 0))

        # キャスト: 左に役の一覧、右に選択中の役の設定
        self.cast_area = area = ctk.CTkFrame(self, fg_color="transparent")
        area.pack(fill="both", expand=True, padx=12, pady=(6, 2))
        area.grid_columnconfigure(0, weight=2, minsize=230)
        area.grid_columnconfigure(1, weight=3)
        area.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(area)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        ctk.CTkLabel(left, text=t("cast"), font=ctk.CTkFont(size=13, weight="bold"),
                     anchor="w").pack(fill="x", padx=10, pady=(6, 0))
        self.role_list = ctk.CTkScrollableFrame(left, fg_color="transparent", height=120)
        self.role_list.pack(fill="both", expand=True, padx=4, pady=2)
        btns = ctk.CTkFrame(left, fg_color="transparent")
        btns.pack(fill="x", padx=6, pady=(0, 6))
        for text, cmd in ((t("role_add"), self.add_role), (t("role_duplicate"), self.duplicate_role),
                          (t("role_delete"), self.delete_role), (t("role_make_default"), self.make_default)):
            ctk.CTkButton(btns, text=text, width=10, height=26, command=cmd).pack(
                side="left", fill="x", expand=True, padx=2)

        self.detail = d = ctk.CTkFrame(area)
        d.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self.detail_title = ctk.CTkLabel(d, text="", font=ctk.CTkFont(size=13, weight="bold"), anchor="w")
        self.detail_title.pack(fill="x", padx=12, pady=(6, 0))
        r = w.row(d)
        w.label(r, t("role_name"))
        self.name_var = ctk.StringVar()
        name_entry = ctk.CTkEntry(r, textvariable=self.name_var, width=130)
        name_entry.pack(side="left", padx=(4, 6))
        name_entry.bind("<FocusOut>", lambda e: self._commit_panel())
        name_entry.bind("<Return>", lambda e: self._commit_panel())
        ctk.CTkButton(r, text=t("copy_tag"), width=10, height=26, command=self._copy_tag).pack(side="left")
        r = w.row(d)
        w.label(r, t("subtitle_label"))
        self.label_var = ctk.StringVar()
        label_entry = ctk.CTkEntry(r, textvariable=self.label_var, width=130)
        label_entry.pack(side="left", padx=(4, 12))
        label_entry.bind("<FocusOut>", lambda e: self._commit_panel())
        ctk.CTkLabel(r, text=t("role_color")).pack(side="left")
        self.color_var = ctk.StringVar()
        w.ColorButton(r, self.color_var, width=96, empty_text=t("color_default"),
                      command=self._commit_panel).pack(side="left", padx=(4, 0))
        self.picker = w.VoicePicker(d, on_change=self._on_voice_changed)
        self.speed_var = ctk.DoubleVar(value=1.0)
        self.pitch_var = ctk.DoubleVar(value=0.0)
        self.intonation_var = ctk.DoubleVar(value=1.0)
        self.volume_var = ctk.DoubleVar(value=1.0)
        self.speed_row = w.SliderRow(d, t("speed"), self.speed_var, 0.5, 2.0, 30, command=self._commit_panel)
        self.pitch_row = w.SliderRow(d, t("pitch"), self.pitch_var, -0.15, 0.15, 30, "{:.2f}",
                                     command=self._commit_panel)
        self.intonation_row = w.SliderRow(d, t("intonation"), self.intonation_var, 0.0, 2.0, 20,
                                          command=self._commit_panel)
        self.volume_row = w.SliderRow(d, t("volume"), self.volume_var, 0.0, 2.0, 20, command=self._commit_panel)
        for sr in (self.speed_row, self.pitch_row, self.intonation_row, self.volume_row):
            sr.pack(fill="x", padx=12, pady=w.ROW_PAD)
        self.instructions_row = r = ctk.CTkFrame(d, fg_color="transparent")
        w.label(r, t("instructions"))
        self.instructions_var = ctk.StringVar()
        ins = ctk.CTkEntry(r, textvariable=self.instructions_var, placeholder_text=t("instructions_hint"))
        ins.pack(side="left", fill="x", expand=True, padx=(4, 0))
        ins.bind("<FocusOut>", lambda e: self._commit_panel())

        # 間・文の区切り・辞書
        r = w.row(self, pady=(4, 2))
        w.label(r, t("pauses"), width=70)
        self.pause_var = ctk.DoubleVar(value=0.5)
        self.turn_pause_var = ctk.StringVar(value="")
        self.end_pause_var = ctk.DoubleVar(value=2.0)
        for text, var, ph in ((t("pause_sentence"), self.pause_var, None),
                              (t("turn_pause"), self.turn_pause_var, t("turn_pause_hint")),
                              (t("end_pause_short"), self.end_pause_var, None)):
            ctk.CTkLabel(r, text=text).pack(side="left", padx=(4, 0))
            kw = {"placeholder_text": ph} if ph else {}
            ctk.CTkEntry(r, textvariable=var, width=46, **kw).pack(side="left", padx=(4, 10))
        self.split_sentence_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(r, text=t("split_sentences"), variable=self.split_sentence_var).pack(side="left")
        self.dict_btn = ctk.CTkButton(r, text=t("dictionary_btn"), width=10, command=self.open_dictionary)
        self.dict_btn.pack(side="right")

        # テスト再生
        r = w.row(self, pady=(2, 2))
        w.label(r, t("test"), width=70).pack_configure(anchor="n")
        self.test_textbox = ctk.CTkTextbox(r, height=56, wrap="word")
        self.test_textbox.insert("1.0", t("test_text"))
        self.test_textbox.pack(side="left", fill="x", expand=True, padx=(4, 6))
        self.test_play_btn = ctk.CTkButton(r, text=t("play"), width=80, command=self.test_play, state="disabled")
        self.test_play_btn.pack(side="left", anchor="n")

        # VOICEVOX の利用規約 (VOICEVOX のみ)
        self.terms_row = r = ctk.CTkFrame(self, fg_color="transparent")
        note = w.hint(r, t("voicevox_terms"), cursor="hand2")
        note.pack(side="left", padx=(74, 0))
        note.bind("<Button-1>", lambda e: webbrowser.open("https://voicevox.hiroshiba.jp/"))

        self._test_stop = False
        self._apply_engine_ui()
        self._load_panel()
        self.refresh_roles()

    def _apply_engine_ui(self):
        voicevox = self.engine == "voicevox"
        self.engine_selector.set(engine_label(self.engine))
        self.url_label.configure(text="URL")
        self.fetch_btn.configure(text=t("fetch_speakers") if voicevox else t("fetch_voices"))
        if voicevox:
            self.openai_row.pack_forget()
            self.instructions_row.pack_forget()
            self.pitch_row.pack(fill="x", padx=12, pady=w.ROW_PAD, before=self.volume_row)
            self.intonation_row.pack(fill="x", padx=12, pady=w.ROW_PAD, before=self.volume_row)
            self.terms_row.pack(fill="x", padx=12, pady=(0, 4))
        else:
            self.openai_row.pack(fill="x", padx=12, pady=w.ROW_PAD, after=self.engine_selector.master)
            self.pitch_row.pack_forget()
            self.intonation_row.pack_forget()
            self.instructions_row.pack(fill="x", padx=12, pady=w.ROW_PAD, after=self.volume_row)
            self.terms_row.pack_forget()

    # ------------------------------------------------------------------
    # 設定の読み書き
    # ------------------------------------------------------------------

    def load(self, s: Settings):
        """Settings の値を画面に反映する (既定の役は <config> の話者、ほかの役は <cast>)。"""
        self._loading = True
        self.set_engine(s.engine)
        self.urls = dict(s.urls)
        self.url_var.set(s.url)
        self.model_var.set(s.model)
        self.api_key_var.set(s.api_key)
        default = Role(pipeline.default_role_name(s), speaker=s.speaker, style=s.style, speed=s.speed,
                       pitch=s.pitch, intonation=s.intonation, volume=s.volume, color=s.role_color,
                       label=s.role_label)
        self.roles = [default] + [copy.deepcopy(r) for r in s.cast.values() if r.name != default.name or not
                                  default.name]
        self.default_idx = 0
        self.selected = 0
        self.pause_var.set(s.pause)
        self.turn_pause_var.set("" if s.turn_pause is None else f"{s.turn_pause:g}")
        self.end_pause_var.set(s.end_pause)
        self.split_sentence_var.set(s.split_sentences)
        self.dictionary = dict(s.dictionary)
        self._loading = False
        self._load_panel()
        self.refresh_roles()

    def save(self, s: Settings):
        """画面の値を Settings に書き込む。"""
        self._commit_panel()
        s.engine = self.engine
        s.urls = dict(self.urls)
        s.urls[self.engine] = self.url_var.get().strip()
        s.model = self.model_var.get().strip()
        s.api_key = self.api_key_var.get().strip()
        default = self.roles[self.default_idx]
        s.speaker, s.style = default.speaker, default.style
        s.speed, s.pitch, s.intonation, s.volume = default.speed, default.pitch, default.intonation, default.volume
        s.role_color, s.role_label = default.color, default.label
        # 役名が話者名と同じなら書かない (<config speaker=...> だけの以前の形のまま)
        s.role = ""
        s.role = default.name if default.name != pipeline.default_role_name(s) else ""
        s.cast = {r.name: copy.deepcopy(r) for i, r in enumerate(self.roles) if i != self.default_idx}
        s.pause = self._float(self.pause_var, 0.5)
        turn = self.turn_pause_var.get().strip()
        try:
            s.turn_pause = float(turn) if turn else None
        except ValueError:
            s.turn_pause = None
        s.end_pause = self._float(self.end_pause_var, 2.0)
        s.split_sentences = self.split_sentence_var.get()
        s.dictionary = dict(self.dictionary)

    @staticmethod
    def _float(var, default: float) -> float:
        try:
            return float(var.get())
        except (ValueError, Exception):
            return default

    # ------------------------------------------------------------------
    # エンジンと話者一覧
    # ------------------------------------------------------------------

    def _on_engine_selected(self, label_text: str):
        engine = next(e for e in ENGINES if engine_label(e) == label_text)
        if engine != self.engine:
            self.set_engine(engine)
            # 文末で区切るかはエンジンの想定言語に合わせる (<config> の split_sentences で上書き可)
            self.split_sentence_var.set(engine != "voicevox")

    def set_engine(self, engine: str):
        """エンジンを切り替える。話者一覧は取り直しになり、役の話者は選び直しになる。"""
        if engine not in ENGINES:
            return
        self.engine_selector.set(engine_label(engine))
        if engine == self.engine:
            return
        self.urls[self.engine] = self.url_var.get().strip()
        self.engine = engine
        self.url_var.set(self.urls.get(engine) or DEFAULT_URLS[engine])
        self.speakers_cache = []
        self.voices = []
        for r in self.roles:
            r.speaker, r.style = "", ""
        self.picker.clear(engine)
        self._apply_engine_ui()
        self._load_panel()
        self.refresh_roles()
        self.app._update_run_btn()

    @property
    def has_voices(self) -> bool:
        return self.picker.fetched

    def speakers_for_cast(self):
        return self.speakers_cache if self.engine == "voicevox" else self.voices

    def fetch(self):
        """話者一覧を別スレッドで取得する (取得中も GUI が固まらないように)。"""
        self._commit_panel()
        self.app._log_clear()
        url = self.url_var.get().strip().rstrip("/")
        name = t("engine_name_voicevox") if self.engine == "voicevox" else t("engine_name_server")
        self.app._log(t("log_connecting", name=name, url=url))
        self.fetch_btn.configure(text=t("fetching"), state="disabled")
        settings = Settings()
        settings.set_engine(self.engine)
        settings.url = url
        settings.api_key = self.api_key_var.get().strip()
        threading.Thread(target=self._fetch_worker, args=(settings,), daemon=True).start()

    def _fetch_worker(self, settings: Settings):
        url = settings.url
        try:
            speakers = pipeline.fetch_speakers(settings)
        except requests.exceptions.ConnectTimeout:
            msg = t("log_connect_timeout", url=url, sec=pipeline.FETCH_TIMEOUT[0])
        except requests.exceptions.ConnectionError:
            msg = t("log_connect_failed", url=url)
        except requests.exceptions.Timeout:
            msg = t("log_read_timeout", url=url, sec=pipeline.FETCH_TIMEOUT[1])
        except Exception as e:
            msg = t("log_fetch_failed", e=e)
        else:
            self.after(0, self.on_fetched, settings.engine, speakers)
            return
        name = t("engine_name_voicevox") if settings.engine == "voicevox" else t("engine_name_server")
        self.after(0, self._on_fetch_failed, msg + t("log_check_server", name=name))

    def _reset_fetch_btn(self):
        self.fetch_btn.configure(text=t("fetch_speakers") if self.engine == "voicevox" else t("fetch_voices"),
                                 state="normal")

    def _on_fetch_failed(self, msg: str):
        self._reset_fetch_btn()
        self.app._log(msg)

    def on_fetched(self, engine: str, speakers, quiet: bool = False):
        """取得した話者一覧を反映する。quiet なら結果をログに出さない (画面の作り直し時)。"""
        self._reset_fetch_btn()
        if engine != self.engine:
            return  # 取得中にエンジンが切り替えられた
        if engine == "voicevox":
            self.speakers_cache = speakers
            self.picker.set_voicevox(speakers)
            if not quiet:
                if self.picker.fetched:
                    n_styles = sum(len(sp.get("styles", [])) for sp in speakers)
                    self.app._log(t("log_speakers_fetched", n=len(speakers), styles=n_styles))
                else:
                    self.app._log(t("log_no_speakers"))
        else:
            self.voices = list(speakers)
            self.picker.set_openai(self.voices)
            if not quiet:
                self.app._log(t("log_voices_fetched", n=len(speakers)) if speakers else t("log_no_voices"))
        # 役の話者を一覧から選び直す (見つからなければ一覧の最初)
        if self.picker.fetched:
            for i, r in enumerate(self.roles):
                self.picker.select(r.speaker, r.style)
                self._set_voice(r, *self.picker.get())
            self._load_panel()
        self.refresh_roles()
        self.app._update_run_btn()
        self.test_play_reset()

    # ------------------------------------------------------------------
    # キャスト
    # ------------------------------------------------------------------

    def _auto_name(self, speaker: str) -> str:
        """話者から決まる役名の既定値 (VOICEVOX: 話者名 / OpenAI互換: 声の名前の部分)"""
        if not speaker:
            return ""
        return voice_short_name(speaker) if self.engine == "openai" else speaker

    def _set_voice(self, role: Role, speaker: str, style: str):
        """役の話者を変える。役名が話者から決まる既定の名前のままで、ノートで使われていなければ、役名も変える。"""
        if speaker != role.speaker and (not role.name or (
                role.name == self._auto_name(role.speaker) and not self.usage.get(role.name))):
            new_name = self._unique_name(self._auto_name(speaker), exclude=role)
            if new_name:
                role.name = new_name
        role.speaker, role.style = speaker, style

    def _unique_name(self, base: str, exclude: Role | None = None) -> str:
        names = {r.name for r in self.roles if r is not exclude}
        if base not in names:
            return base
        n = 2
        while f"{base}{n}" in names:
            n += 1
        return f"{base}{n}"

    def _on_voice_changed(self):
        if self._loading:
            return
        role = self.roles[self.selected]
        self._set_voice(role, *self.picker.get())
        self.name_var.set(role.name)
        self.refresh_roles()

    def _load_panel(self):
        """選択中の役を右の設定欄に出す。"""
        self._loading = True
        role = self.roles[self.selected]
        self.detail_title.configure(text=t("role_settings", name=role.name or "-"))
        self.name_var.set(role.name)
        self.label_var.set(role.label)
        self.color_var.set(role.color)
        self.picker.select(role.speaker, role.style)
        if self.picker.fetched:
            role.speaker, role.style = self.picker.get()
        self.speed_var.set(role.speed)
        self.pitch_var.set(role.pitch)
        self.intonation_var.set(role.intonation)
        self.volume_var.set(role.volume)
        self.instructions_var.set(role.instructions)
        self._loading = False

    def _commit_panel(self):
        """右の設定欄の値を、選択中の役に書き込む。"""
        if self._loading or not self.roles:
            return
        role = self.roles[self.selected]
        name = self.name_var.get().strip()
        if name != role.name:
            if not name or _BAD_NAME.search(name) or any(r.name == name for r in self.roles if r is not role):
                self.app._log(t("log_role_name_invalid", name=name))
                self.name_var.set(role.name)
            else:
                if self.usage.get(role.name):
                    self.app._log(t("log_role_renamed_used", old=role.name, n=self.usage[role.name]))
                role.name = name
        role.label = self.label_var.get().strip()
        role.color = self.color_var.get()
        if self.picker.fetched:
            role.speaker, role.style = self.picker.get()
        role.speed = round(self.speed_var.get(), 2)
        role.pitch = round(self.pitch_var.get(), 3)
        role.intonation = round(self.intonation_var.get(), 2)
        role.volume = round(self.volume_var.get(), 2)
        role.instructions = self.instructions_var.get().strip()
        self.detail_title.configure(text=t("role_settings", name=role.name or "-"))
        self.refresh_roles()

    def select_role(self, index: int):
        self._commit_panel()
        self.selected = index
        self._load_panel()
        self.refresh_roles()

    def set_usage(self, usage: dict[str | None, int]):
        """入力ファイルのノートで、役ごとに何文あるか (None は <speaker> のない文)"""
        self.usage = dict(usage)
        self.refresh_roles()

    def role_usage(self, index: int) -> int:
        n = self.usage.get(self.roles[index].name, 0)
        if index == self.default_idx:
            n += self.usage.get(None, 0)
        return n

    def undefined_roles(self) -> list[str]:
        names = {r.name for r in self.roles}
        return [name for name in self.usage if name is not None and name not in names]

    def refresh_roles(self):
        """役の一覧を作り直す (続けて呼ばれても、作り直すのは手が空いたときに1回だけ)。"""
        if not getattr(self, "_refresh_pending", False):
            self._refresh_pending = True
            self.after_idle(self._refresh_roles_now)

    def _refresh_roles_now(self):
        self._refresh_pending = False
        if not self.winfo_exists():
            return
        for child in self.role_list.winfo_children():
            child.destroy()
        selected_color = ("#E0E7FF", "#3B3F6B")
        for i, role in enumerate(self.roles):
            rf = ctk.CTkFrame(self.role_list, fg_color=selected_color if i == self.selected else "transparent",
                              corner_radius=6)
            rf.pack(fill="x", pady=1)
            swatch = ctk.CTkFrame(rf, width=14, height=14, corner_radius=3, border_width=1,
                                  border_color=("gray60", "gray40"),
                                  fg_color=role.color or self.app.default_font_color())
            swatch.pack(side="left", padx=(6, 4))
            count = self.role_usage(i)
            name = ctk.CTkLabel(rf, text=role.name or "-", anchor="w",
                                text_color=None if count or not self.usage else w.HINT_COLOR)
            name.pack(side="left", fill="x", expand=True)
            if i == self.default_idx:
                w.hint(rf, t("role_default_mark")).pack(side="left", padx=4)
            w.hint(rf, t("role_count", n=count) if self.usage else "").pack(side="left", padx=4)
            ctk.CTkButton(rf, text="▶", width=26, height=22, command=lambda i=i: self.play_role(i)).pack(
                side="left", padx=(2, 4), pady=2)
            for widget in (rf, swatch, name):
                widget.bind("<Button-1>", lambda e, i=i: self.select_role(i))
        for name in self.undefined_roles():
            rf = ctk.CTkFrame(self.role_list, fg_color="transparent")
            rf.pack(fill="x", pady=1)
            ctk.CTkLabel(rf, text=t("role_undefined", name=name), anchor="w",
                         text_color=("#B45309", "#FBBF24")).pack(side="left", fill="x", expand=True, padx=(6, 0))
            w.hint(rf, t("role_count", n=self.usage[name])).pack(side="left", padx=4)
            ctk.CTkButton(rf, text=t("role_define"), width=10, height=22,
                          command=lambda n=name: self.add_role(n)).pack(side="left", padx=(2, 4), pady=2)

    def _next_voice(self) -> tuple[str, str]:
        """一覧で、最後の役の次の話者"""
        speakers = self.picker.speakers()
        if not speakers:
            return "", ""
        last = self.roles[-1].speaker
        i = speakers.index(last) + 1 if last in speakers else 0
        return speakers[i % len(speakers)], ""

    def add_role(self, name: str | None = None):
        self._commit_panel()
        speaker, style = self._next_voice()
        if name is None:
            name = self._unique_name(self._auto_name(speaker) or t("role_new_name", n=len(self.roles) + 1))
        used = {r.color for r in self.roles}
        color = next((c for c in PALETTE if c not in used), PALETTE[len(self.roles) % len(PALETTE)])
        self.roles.append(Role(name, speaker=speaker, style=style, color=color))
        self.selected = len(self.roles) - 1
        self._load_panel()
        self.refresh_roles()

    def duplicate_role(self):
        self._commit_panel()
        role = copy.deepcopy(self.roles[self.selected])
        role.name = self._unique_name(role.name)
        self.roles.append(role)
        self.selected = len(self.roles) - 1
        self._load_panel()
        self.refresh_roles()

    def delete_role(self):
        if self.selected == self.default_idx:
            self.app._log(t("log_cannot_delete_default"))
            return
        role = self.roles[self.selected]
        n = self.usage.get(role.name, 0)
        if n and not messagebox.askyesno(t("role_delete"), t("confirm_delete_role", name=role.name, n=n)):
            return
        del self.roles[self.selected]
        if self.default_idx > self.selected:
            self.default_idx -= 1
        self.selected = self.default_idx
        self._load_panel()
        self.refresh_roles()

    def make_default(self):
        self._commit_panel()
        self.default_idx = self.selected
        self.refresh_roles()

    def _copy_tag(self):
        self._commit_panel()
        tag = f"<speaker={self.roles[self.selected].name}>"
        self.clipboard_clear()
        self.clipboard_append(tag)
        self.app._log(t("log_tag_copied", tag=tag))

    # ------------------------------------------------------------------
    # 辞書
    # ------------------------------------------------------------------

    def open_dictionary(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title(t("dict_title"))
        dialog.geometry("520x460")
        dialog.grab_set()
        w.hint(dialog, t("dict_desc"), justify="left").pack(fill="x", padx=12, pady=(10, 4))
        head = ctk.CTkFrame(dialog, fg_color="transparent")
        head.pack(fill="x", padx=16)
        ctk.CTkLabel(head, text=t("dict_word"), width=180, anchor="w").pack(side="left")
        ctk.CTkLabel(head, text=t("dict_reading"), anchor="w").pack(side="left", padx=(6, 0))
        body = ctk.CTkScrollableFrame(dialog)
        body.pack(fill="both", expand=True, padx=10, pady=4)
        rows: list[tuple[ctk.CTkFrame, ctk.StringVar, ctk.StringVar]] = []

        def add_row(word="", reading=""):
            rf = ctk.CTkFrame(body, fg_color="transparent")
            rf.pack(fill="x", pady=1)
            wv, rv = ctk.StringVar(value=word), ctk.StringVar(value=reading)
            ctk.CTkEntry(rf, textvariable=wv, width=180).pack(side="left")
            ctk.CTkEntry(rf, textvariable=rv).pack(side="left", fill="x", expand=True, padx=6)
            item = (rf, wv, rv)
            ctk.CTkButton(rf, text="×", width=28, command=lambda: (rows.remove(item), rf.destroy())).pack(
                side="left")
            rows.append(item)

        for word, reading in self.dictionary.items():
            add_row(word, reading)
        if not self.dictionary:
            add_row()

        def on_ok():
            self.dictionary = {wv.get().strip(): rv.get().strip() for _, wv, rv in rows
                               if wv.get().strip() and rv.get().strip()}
            dialog.destroy()

        foot = ctk.CTkFrame(dialog, fg_color="transparent")
        foot.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(foot, text=t("dict_add"), width=90, command=add_row).pack(side="left")
        ctk.CTkButton(foot, text="OK", width=120, command=on_ok).pack(side="right")

    # ------------------------------------------------------------------
    # テスト再生
    # ------------------------------------------------------------------

    def test_play(self):
        if self.test_play_btn.cget("text") == t("stop_play"):
            self._test_stop = True
            winsound.PlaySound(None, winsound.SND_PURGE)
            self.test_play_reset()
            return
        text = self.test_textbox.get("1.0", "end-1c").strip()
        if text:
            self._play(text)

    def play_role(self, index: int):
        """その役の声で、テスト欄の1行目 (タグを除く) を読む。"""
        self._commit_panel()
        first = next((ln for ln in self.test_textbox.get("1.0", "end-1c").splitlines() if ln.strip()), "")
        sample = _SPEAKER_TAGS.sub("", first).strip() or t("test_text")
        self._play(f"<speaker={self.roles[index].name}>{sample}")

    def _play(self, text: str):
        if not self.has_voices:
            self.app._log(t("log_fetch_first"))
            return
        if self.test_play_btn.cget("text") != t("play"):
            return
        settings = self.app._settings_from_ui()
        engine, cast = self.app._create_engine(settings)
        self.test_play_btn.configure(text=t("synthesizing"), state="disabled")
        threading.Thread(target=self._test_play_worker, args=(text, settings, engine, cast), daemon=True).start()

    def _test_play_worker(self, text, settings, engine, cast):
        try:
            wav, timings, _ = pipeline.synthesize_notes(engine, text, cast, settings.pause,
                                                        settings.split_sentences, turn_pause=settings.turn_pause,
                                                        dictionary=settings.dictionary)
            self.after(0, lambda: self.test_play_btn.configure(text=t("stop_play"), state="normal"))
            self._test_stop = False
            if timings:
                self.after(0, lambda: self.app._log(t("log_subtitle_preview")))
                threading.Thread(target=self._test_subtitle_worker, args=(timings, len(cast.used) > 1),
                                 daemon=True).start()
            # SND_MEMORY は同期再生 (再生完了 or SND_PURGE で停止するまでブロック)
            winsound.PlaySound(wav, winsound.SND_MEMORY)
            self._test_stop = True
        except Exception as e:
            self._test_stop = True
            self.after(0, lambda msg=str(e): self.app._log(t("log_test_error", msg=msg)))
        finally:
            self.after(0, self.test_play_reset)

    _STRIP_TAGS = re.compile(r"</?[a-zA-Z][^>]*>")

    def _test_subtitle_worker(self, timings, show_role: bool):
        """再生タイミングに合わせて字幕テキストをログに表示する。"""
        import time
        t0 = time.perf_counter()
        for timing in timings:
            if self._test_stop:
                return
            wait = timing.start_ms / 1000.0 - (time.perf_counter() - t0)
            if wait > 0:
                time.sleep(wait)
            if self._test_stop:
                return
            clean = pipeline._unescape_display(self._STRIP_TAGS.sub("", timing.text))
            prefix = f"{timing.speaker}: " if show_role and timing.speaker else ""
            self.after(0, lambda s=f"  {prefix}{clean}\n": self.app._log(s))

    def test_play_reset(self):
        self.test_play_btn.configure(text=t("play"), state="normal" if self.has_voices else "disabled")
