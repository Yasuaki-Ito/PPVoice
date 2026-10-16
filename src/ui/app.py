"""PPVoice のメインウィンドウ: 上の帯 (ファイル・生成) + 3つのタブ (音声・字幕・その他) + ログ"""

import os
import sys
import threading
import webbrowser
from tkinter import filedialog, messagebox

import customtkinter as ctk
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except ImportError:
    _HAS_DND = False

import hashlib
import tempfile

from PIL import Image

import i18n
import pipeline
import powerpoint
import update_check
from cache import SynthCache
from i18n import t
from lint import lint_notes
from notes import parse_notes
from pptx_reader import read_slides
from settings import Settings
from ui.other_tab import OtherTab
from ui.subtitle_tab import SubtitleTab
from ui.voice_tab import VoiceTab
from version import __version__

# スライドの選択のダイアログのサムネイルの幅
THUMB_WIDTH = 120

# ドキュメントサイト・リポジトリ
DOCS_URL = "https://yasuaki-ito.github.io/PPVoice/"
DOCS_URL_EN = "https://yasuaki-ito.github.io/PPVoice/en/"
GITHUB_URL = "https://github.com/Yasuaki-Ito/PPVoice"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme(os.path.join(os.path.dirname(os.path.dirname(__file__)), "theme_modern.json"))


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
    def __init__(self):
        super().__init__()
        self.title(f"PPVoice v{__version__}")
        self.geometry("1100x720")
        self.minsize(960, 600)
        # アイコン (src/ 内 → ルート の順で探す)
        src = os.path.dirname(os.path.dirname(__file__))
        for d in (src, os.path.join(src, "..")):
            ico_path = os.path.join(d, "app.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
                break

        i18n.load_language()
        self._running = False
        self._cancel_event = threading.Event()
        self._selected_slides: set[int] | None = None  # None = 全スライド

        self._build_ui()
        self._setup_dnd()
        self._update_run_btn()
        self._start_update_check()

    # ------------------------------------------------------------------
    # 画面
    # ------------------------------------------------------------------

    def _build_ui(self):
        # (言語の切り替え時は container ごと作り直す → _rebuild_ui)
        self._container = c = ctk.CTkFrame(self, fg_color="transparent")
        c.pack(fill="both", expand=True, padx=12, pady=12)
        c.grid_columnconfigure(0, weight=3, minsize=600)
        c.grid_columnconfigure(1, weight=2, minsize=280)
        c.grid_rowconfigure(1, weight=1)

        self._build_top_band(c)

        # 設定のタブ
        self.tabs = ctk.CTkTabview(c, anchor="nw")
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(8, 0))
        self._tab_names = [t("tab_voice"), t("tab_subtitle"), t("tab_other")]
        self.voice_tab = VoiceTab(self.tabs.add(self._tab_names[0]), self)
        self.voice_tab.pack(fill="both", expand=True)
        self.subtitle_tab = SubtitleTab(self.tabs.add(self._tab_names[1]), self)
        self.subtitle_tab.pack(fill="both", expand=True)
        self.other_tab = OtherTab(self.tabs.add(self._tab_names[2]), self)
        self.other_tab.pack(fill="both", expand=True)
        self.voice_tab.refresh_roles()

        self._build_log(c)
        self.input_var.trace_add("write", lambda *_: self._update_run_btn())
        self.apply_powerpoint_ui()

    def _build_top_band(self, parent):
        band = ctk.CTkFrame(parent)
        band.grid(row=0, column=0, columnspan=2, sticky="ew")
        band.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(band, text=t("input_short")).grid(row=0, column=0, sticky="w", padx=(12, 6), pady=(10, 3))
        self.input_var = ctk.StringVar()
        ctk.CTkEntry(band, textvariable=self.input_var).grid(row=0, column=1, sticky="ew", pady=(10, 3))
        ctk.CTkButton(band, text=t("browse"), width=60, command=self._browse_input).grid(
            row=0, column=2, padx=6, pady=(10, 3))
        slides = ctk.CTkFrame(band, fg_color="transparent")
        slides.grid(row=0, column=3, sticky="w", padx=(6, 12), pady=(10, 3))
        ctk.CTkLabel(slides, text=t("slides")).pack(side="left", padx=(0, 8))
        self.slide_range_var = ctk.StringVar(value="all")
        for text, value in ((t("slides_all"), "all"), (t("slides_some"), "select")):
            ctk.CTkRadioButton(slides, text=text, variable=self.slide_range_var, value=value, width=10,
                               command=self._on_slide_range_changed).pack(side="left", padx=(0, 10))
        self.slide_select_btn = ctk.CTkButton(slides, text=t("slides_select"), width=60,
                                              command=self._open_slide_selector)
        self.slide_select_label = ctk.CTkLabel(slides, text="")

        ctk.CTkLabel(band, text=t("output_short")).grid(row=1, column=0, sticky="w", padx=(12, 6), pady=(3, 10))
        self.output_var = ctk.StringVar()
        ctk.CTkEntry(band, textvariable=self.output_var).grid(row=1, column=1, sticky="ew", pady=(3, 10))
        ctk.CTkButton(band, text=t("browse"), width=60, command=self._browse_output).grid(
            row=1, column=2, padx=6, pady=(3, 10))
        run = ctk.CTkFrame(band, fg_color="transparent")
        run.grid(row=1, column=3, sticky="ew", padx=(6, 12), pady=(3, 10))
        self.video_var = ctk.BooleanVar(value=False)
        self.video_check = ctk.CTkCheckBox(run, text=t("make_video"), variable=self.video_var, width=10)
        self.video_check.pack(side="left", padx=(0, 10))
        self.run_btn = ctk.CTkButton(run, text=t("run"), font=ctk.CTkFont(size=15, weight="bold"),
                                     height=36, width=170, corner_radius=12, command=self._on_run)
        self.run_btn.pack(side="left", fill="x", expand=True)

    def _build_log(self, parent):
        sec = ctk.CTkFrame(parent)
        sec.grid(row=1, column=1, sticky="nsew", padx=(6, 0), pady=(8, 0))
        # ドキュメント・GitHub へのリンクと表示言語
        head = ctk.CTkFrame(sec, fg_color="transparent")
        head.pack(fill="x", padx=12, pady=(8, 4))
        for text, url in ((t("link_docs"), DOCS_URL_EN if i18n.get_language() == "en" else DOCS_URL),
                          ("GitHub", GITHUB_URL)):
            link = ctk.CTkLabel(head, text=text, cursor="hand2", text_color=("#4F46E5", "#A5B4FC"),
                                font=ctk.CTkFont(underline=True))
            link.pack(side="left", padx=(0, 12))
            link.bind("<Button-1>", lambda e, u=url: webbrowser.open(u))
        self.language_selector = ctk.CTkSegmentedButton(head, values=list(i18n.LANGUAGES.values()),
                                                        command=self._on_language_selected)
        self.language_selector.set(i18n.LANGUAGES[i18n.get_language()])
        self.language_selector.pack(side="right")
        self.progress = ctk.CTkProgressBar(sec, height=6, corner_radius=3)
        self.progress.set(0)
        # 生成の後に、出力をスライドショーで確認する (PowerPoint と連携できるときだけ)
        self.slideshow_btn = ctk.CTkButton(sec, text=t("slideshow"), height=30, command=self._start_slideshow)
        self._last_output: str | None = None
        # 進捗は生成中だけ表示する
        self.log_box = ctk.CTkTextbox(sec, state="disabled", font=ctk.CTkFont(size=12), wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    # --- 回帰確認・画面キャプチャのスクリプトから使う名前 ---
    @property
    def fetch_btn(self):
        return self.voice_tab.fetch_btn

    def _fetch_speakers(self):
        self.voice_tab.fetch()

    # ------------------------------------------------------------------
    # 言語の切り替え
    # ------------------------------------------------------------------

    def _on_language_selected(self, label_text: str):
        lang = next(k for k, v in i18n.LANGUAGES.items() if v == label_text)
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
        vt = self.voice_tab
        state = {
            "input": self.input_var.get(), "output": self.output_var.get(),
            "slide_range": self.slide_range_var.get(), "selected": self._selected_slides,
            "test_text": vt.test_textbox.get("1.0", "end-1c"),
            "log": self.log_box.get("1.0", "end-1c"),
            "speakers": vt.speakers_cache if vt.engine == "voicevox" else vt.voices,
            "usage": vt.usage,
            "tab": self._tab_names.index(self.tabs.get()),
        }
        self._container.destroy()
        self._build_ui()

        self.input_var.set(state["input"])
        self.output_var.set(state["output"])
        self.slide_range_var.set(state["slide_range"])
        self._on_slide_range_changed()
        self._selected_slides = state["selected"]
        self._load_settings_to_ui(settings)
        self.voice_tab.usage = state["usage"]
        if state["speakers"]:
            self.voice_tab.on_fetched(settings.engine, state["speakers"], quiet=True)
        self.voice_tab.refresh_roles()
        # テスト欄が初期値のままなら、新しい言語の初期値にする
        if state["test_text"] not in i18n.STRINGS["test_text"]:
            self.voice_tab.test_textbox.delete("1.0", "end")
            self.voice_tab.test_textbox.insert("1.0", state["test_text"])
        self.tabs.set(self._tab_names[state["tab"]])
        if state["log"]:
            self._log(state["log"] + "\n")
        self._update_run_btn()

    # ------------------------------------------------------------------
    # 設定
    # ------------------------------------------------------------------

    def _settings_from_ui(self) -> Settings:
        """画面の値から Settings を作る。"""
        s = Settings()
        self.voice_tab.save(s)
        self.subtitle_tab.save(s)
        self.other_tab.save(s)
        return s

    def _load_settings_to_ui(self, s: Settings):
        """Settings の値を画面に反映する。"""
        self.subtitle_tab.load(s)
        self.other_tab.load(s)
        self.voice_tab.load(s)

    def _apply_config(self, config: dict, casts: dict | None = None):
        """解析済みの <config> / <cast> を画面に反映する。"""
        s = self._settings_from_ui()
        s.apply_config(config, warn=lambda k, v: self._log(t("log_config_invalid", key=k, value=v)))
        for name, values in (casts or {}).items():
            s.apply_cast({name: values}, warn=lambda k, v, name=name: self._log(
                t("log_cast_invalid", name=name, key=k, value=v)))
        self._load_settings_to_ui(s)

    def powerpoint_enabled(self) -> bool:
        """PowerPoint と連携するか (PowerPoint があり、その他タブでオフにしていない)"""
        return powerpoint.is_available() and bool(i18n.get_user_setting("use_powerpoint", True))

    def apply_powerpoint_ui(self):
        """PowerPoint と連携できないときは、連携の機能 (動画・プレビュー・サムネイル・スライドショー) を隠す。"""
        enabled = self.powerpoint_enabled()
        if enabled:
            self.video_check.pack(side="left", padx=(0, 10), before=self.run_btn)
        else:
            self.video_check.pack_forget()
            self.video_var.set(False)
            self.slideshow_btn.pack_forget()
        self.subtitle_tab.set_preview_enabled(enabled)

    def default_font_color(self) -> str:
        tab = getattr(self, "subtitle_tab", None)
        return tab.font_color_var.get() if tab else "#FFFFFF"

    def on_font_color_changed(self):
        self.voice_tab.refresh_roles()

    def _create_engine(self, settings: Settings):
        """設定から TTS エンジンと役の声を作る (取得済みの話者一覧を使う)。"""
        return pipeline.create_engine(settings), pipeline.create_cast(
            settings, speakers=self.voice_tab.speakers_for_cast())

    def show_config_tags(self):
        """<config> / <cast> / <dict> タグを表示してコピーできるようにする。"""
        s = self._settings_from_ui()
        tag = "\n".join(x for x in [s.to_config_tag(), *s.to_cast_tags(), s.to_dict_tag()] if x)
        dialog = ctk.CTkToplevel(self)
        dialog.title(t("save_config"))
        dialog.geometry("640x300")
        dialog.grab_set()
        ctk.CTkLabel(dialog, text=t("save_config_desc"), font=ctk.CTkFont(size=12), justify="left").pack(
            padx=16, pady=(16, 8), anchor="w")
        box = ctk.CTkTextbox(dialog, font=ctk.CTkFont(size=11), wrap="word")
        box.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        box.insert("1.0", tag)
        box.configure(state="disabled")

        def on_copy():
            self.clipboard_clear()
            self.clipboard_append(tag)
            copy_btn.configure(text=t("copied"))
            dialog.after(1500, lambda: copy_btn.configure(text=t("copy")))

        copy_btn = ctk.CTkButton(dialog, text=t("copy"), width=120, command=on_copy)
        copy_btn.pack(pady=(0, 16))

    # ------------------------------------------------------------------
    # ファイル
    # ------------------------------------------------------------------

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
        try:
            files = self.tk.splitlist(event.data)
        except Exception:
            files = [event.data]
        for f in files:
            f = f.strip("{}")
            if f.lower().endswith(".pptx"):
                self._set_input_file(f)
                break

    def _browse_input(self):
        path = filedialog.askopenfilename(filetypes=[("PowerPoint", "*.pptx")])
        if path:
            self._set_input_file(path)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(filetypes=[("PowerPoint", "*.pptx")], defaultextension=".pptx")
        if path:
            self.output_var.set(path)

    def _set_input_file(self, path: str):
        """入力ファイルを設定し、<config> / <cast> / <dict> タグを読み込む (役と辞書はファイルごとに読み直す)。"""
        self.input_var.set(path)
        self.output_var.set(os.path.splitext(path)[0] + "_speech.pptx")
        try:
            config = pipeline.read_config(path)
            casts = pipeline.read_cast(path)
            dictionary = pipeline.read_dict(path)
            s = self._settings_from_ui()
            s.cast = {}
            s.dictionary = dictionary
            self._load_settings_to_ui(s)
            if config or casts or dictionary:
                self._apply_config(config, casts)
                details = "\n".join([f"  {k}={v}" for k, v in config.items()]
                                    + [f"  <cast {name}>" for name in casts]
                                    + [f"  <dict> {w_}={r}" for w_, r in dictionary.items()])
                self._log(t("log_config_loaded", details=details))
            self._count_roles(path)
            self.subtitle_tab.set_slides(path, [si.index + 1 for si in read_slides(path) if si.notes_text])
        except Exception:
            pass

    def _count_roles(self, path: str):
        """ノートで役ごとに何文あるかを数え、役の一覧に出す。"""
        s = self._settings_from_ui()
        usage: dict[str | None, int] = {}
        for si in read_slides(path):
            if si.notes_text:
                for sentence in parse_notes(si.notes_text, s.split_sentences, s.dictionary).sentences:
                    usage[sentence.speaker] = usage.get(sentence.speaker, 0) + 1
        self.voice_tab.set_usage(usage)

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

        dialog = ctk.CTkToplevel(self)
        dialog.title(t("slide_select_title"))
        dialog.geometry("360x420")
        dialog.resizable(False, True)
        dialog.grab_set()
        ctk.CTkLabel(dialog, text=t("slide_select_heading", total=total),
                     font=ctk.CTkFont(size=14, weight="bold")).pack(padx=10, pady=(10, 4))
        btn_row = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_row.pack(fill="x", padx=10, pady=(0, 4))
        check_vars: list[ctk.BooleanVar] = []
        ctk.CTkButton(btn_row, text=t("select_all"), width=70,
                      command=lambda: [v.set(True) for v in check_vars]).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_row, text=t("deselect_all"), width=70,
                      command=lambda: [v.set(False) for v in check_vars]).pack(side="left")
        scroll = ctk.CTkScrollableFrame(dialog)
        scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        prev = self._selected_slides
        thumbs: dict[int, ctk.CTkLabel] = {}
        for i in range(total):
            num = i + 1
            var = ctk.BooleanVar(value=(prev is None or num in prev))
            check_vars.append(var)
            preview = (slides[i].notes_text or "").replace("\n", " ")[:30]
            text = t("slide_n", num=num) + (f" - {preview}" if preview else "")
            r = ctk.CTkFrame(scroll, fg_color="transparent")
            r.pack(fill="x", pady=1)
            if self.powerpoint_enabled():
                frame = ctk.CTkFrame(r, border_width=1, border_color=("gray70", "gray40"), fg_color="transparent", corner_radius=0)
                frame.pack(side="left", padx=(0, 8))
                thumbs[num] = ctk.CTkLabel(frame, text="", width=THUMB_WIDTH, height=THUMB_WIDTH * 9 // 16)
                thumbs[num].pack(padx=1, pady=1)
            ctk.CTkCheckBox(r, text=text, variable=var).pack(side="left", anchor="w")
        if thumbs:
            self._load_thumbnails(input_path, total, thumbs, dialog)

        def on_ok():
            selected = {i + 1 for i, v in enumerate(check_vars) if v.get()}
            if not selected:
                messagebox.showwarning(t("no_selection_title"), t("no_selection"), parent=dialog)
                return
            self._selected_slides = selected
            self._update_slide_label(total)
            dialog.destroy()

        ctk.CTkButton(dialog, text="OK", width=120, command=on_ok).pack(pady=(0, 10))

    def _load_thumbnails(self, path: str, total: int, labels: dict, dialog):
        """スライドのサムネイルを PowerPoint で作り (入力ファイルの更新日時ごとにキャッシュ)、ダイアログに出す。"""
        key = hashlib.sha1(f"{os.path.abspath(path)}|{os.path.getmtime(path)}".encode()).hexdigest()[:16]
        out = os.path.join(tempfile.gettempdir(), "ppvoice_thumbs", key)

        def show(paths):
            if not dialog.winfo_exists():
                return
            for num, p in paths.items():
                if num in labels and os.path.exists(p):
                    img = Image.open(p)
                    h = THUMB_WIDTH * img.height // img.width
                    labels[num].configure(image=ctk.CTkImage(img, size=(THUMB_WIDTH, h)))

        def worker():
            paths = {n: os.path.join(out, f"slide{n}.png") for n in range(1, total + 1)}
            missing = [n for n, p in paths.items() if not os.path.exists(p)]
            try:
                if missing:
                    powerpoint.export_slides(path, missing, out, width=THUMB_WIDTH * 2)
            except powerpoint.PowerPointError as e:
                self.after(0, lambda: self._log(t("log_powerpoint_failed", e=e)))
                return
            self.after(0, lambda: show(paths))

        threading.Thread(target=worker, daemon=True).start()

    def _update_slide_label(self, total: int):
        """選択されたスライド番号を表示する ([1,2,3,5,7,8] → "1-3, 5, 7-8")。"""
        sel = self._selected_slides
        if sel is None or len(sel) == total:
            self.slide_select_label.configure(text="")
            return
        nums = sorted(sel)
        parts = []
        start = end = nums[0]
        for n in nums[1:]:
            if n == end + 1:
                end = n
            else:
                parts.append(f"{start}-{end}" if start != end else str(start))
                start = end = n
        parts.append(f"{start}-{end}" if start != end else str(start))
        self.slide_select_label.configure(text=", ".join(parts))

    # ------------------------------------------------------------------
    # ノートの確認 (<next> とアニメーション、書き間違い)
    # ------------------------------------------------------------------

    def check_notes(self):
        path = self.input_var.get()
        if not path or not os.path.isfile(path):
            self._log(t("log_no_input"))
            return
        try:
            next_rows = {r.slide: r for r in pipeline.check_next_tags(path)}
            slides = read_slides(path)
        except Exception as e:
            self._log(t("log_pptx_read_failed", e=e))
            return
        s = self._settings_from_ui()
        vt = self.voice_tab
        vt._commit_panel()
        roles = {r.name for r in vt.roles}
        default = vt.roles[vt.default_idx].name
        styles = None
        if s.engine == "voicevox" and vt.speakers_cache:
            by_speaker = {sp["name"]: {st["name"] for st in sp.get("styles", [])} for sp in vt.speakers_cache}
            styles = {r.name: by_speaker.get(r.speaker, set()) for r in vt.roles}
        self._log(t("log_next_check_header"))
        for r in next_rows.values():
            self._log(t("log_next_check_row", num=r.slide, anims=r.animations, nexts=r.nexts, status=r.status))
        issues = [(si.index + 1, i) for si in slides if si.notes_text
                  for i in lint_notes(si.notes_text, roles, default, styles, s.engine)]
        self._log(t("log_lint_header") + "\n")
        for num, issue in issues:
            self._log(t("log_lint_issue", slide=num, line=issue.line, message=issue.message) + "\n")
        if not issues:
            self._log(t("log_notes_ok"))
        self._log("---\n")

    # ------------------------------------------------------------------
    # 更新の確認
    # ------------------------------------------------------------------

    def _start_update_check(self):
        """新しいバージョンがあればログに知らせる (配布版のみ。ユーザー設定 check_updates でオフにできる)。"""
        if not getattr(sys, "frozen", False) or not i18n.get_user_setting("check_updates", True):
            return

        def worker():
            found = update_check.check(__version__)
            if found:
                self.after(0, lambda: self._log(t("log_update_available", version=found[0], url=found[1])))
        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------
    # ログ
    # ------------------------------------------------------------------

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
    # 生成
    # ------------------------------------------------------------------

    def _update_run_btn(self):
        """入力ファイル・話者一覧の状態に応じてボタンの有効/無効を切り替える。"""
        if self._running or not hasattr(self, "run_btn") or not hasattr(self, "voice_tab"):
            return
        ok = bool(self.input_var.get().strip()) and self.voice_tab.has_voices
        self.run_btn.configure(state="normal" if ok else "disabled")
        if self.voice_tab.test_play_btn.cget("text") == t("play"):
            self.voice_tab.test_play_reset()

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
        output_path = self.output_var.get().strip() or os.path.splitext(input_path)[0] + "_speech.pptx"
        existing = [output_path] if os.path.exists(output_path) else []
        if self.video_var.get() and os.path.exists(self._video_path(output_path)):
            existing.append(self._video_path(output_path))
        if existing:
            names = "\n".join(os.path.basename(f) for f in existing)
            if not messagebox.askyesno(t("overwrite_title"), t("overwrite_msg", names=names)):
                return

        self._running = True
        self._cancel_event.clear()
        self._log_clear()
        self.run_btn.configure(text=t("stop"), fg_color="#EF4444", hover_color="#DC2626")
        self.slideshow_btn.pack_forget()
        self._last_output = None
        self.progress.set(0)
        self.progress.pack(fill="x", padx=12, pady=(0, 8), before=self.log_box)
        threading.Thread(target=self._run_generate, daemon=True).start()

    def _run_generate(self):
        old_stdout = sys.stdout
        sys.stdout = LogRedirector(self.log_box)
        try:
            self._do_generate()
        except pipeline.Cancelled:
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
        self.run_btn.configure(text=t("run"), fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"],
                               hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"])
        self.other_tab.refresh_cache_size()
        if self._last_output and self.powerpoint_enabled():
            self.slideshow_btn.pack(fill="x", padx=12, pady=(0, 8), before=self.log_box)
        self._update_run_btn()

    def _do_generate(self):
        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip() or os.path.splitext(input_path)[0] + "_speech.pptx"
        settings = self._settings_from_ui()
        engine, cast = self._create_engine(settings)
        result = pipeline.generate(
            settings, input_path, output_path,
            selected_slides=self._selected_slides,
            engine=engine, cast=cast, cache=SynthCache() if self.other_tab.use_cache else None,
            on_progress=lambda v: self.after(0, self.progress.set, v),
            should_cancel=self._cancel_event.is_set,
        )
        if not result:
            return
        self._last_output = result.output
        if not self.powerpoint_enabled():
            print(t("log_video_howto"))
        elif self.video_var.get():
            self._export_video(result.output)

    @staticmethod
    def _video_path(output_path: str) -> str:
        return os.path.splitext(output_path)[0] + ".mp4"

    def _export_video(self, pptx: str):
        """生成した PPTX を PowerPoint で動画に書き出す (生成のスレッドから呼ぶ)。"""
        mp4 = self._video_path(pptx)
        print(t("log_video_start", path=mp4))
        self.after(0, lambda: self.progress.configure(mode="indeterminate") or self.progress.start())
        last = [0.0]

        def on_progress(sec: float):
            if sec - last[0] >= 10:
                last[0] = sec
                print(t("log_video_progress", sec=int(sec)))

        try:
            powerpoint.create_video(pptx, mp4, height=int(i18n.get_user_setting("video_height", 1080)),
                                    on_progress=on_progress, should_cancel=self._cancel_event.is_set)
            print(t("log_video_done", path=mp4))
        except powerpoint.PowerPointError as e:
            if self._cancel_event.is_set():
                raise pipeline.Cancelled()
            print(t("log_video_failed", e=e))
        finally:
            self.after(0, lambda: self.progress.stop() or self.progress.configure(mode="determinate"))

    def _start_slideshow(self):
        """生成した PPTX を PowerPoint で開き、スライドショーを始める (選んだスライドの最初から)。"""
        path = self._last_output
        if not path or not os.path.exists(path):
            return
        first = min(self._selected_slides) if self._selected_slides else 1

        def worker():
            try:
                powerpoint.start_slideshow(path, first)
            except powerpoint.PowerPointError as e:
                self.after(0, lambda: self._log(t("log_powerpoint_failed", e=e)))

        threading.Thread(target=worker, daemon=True).start()
