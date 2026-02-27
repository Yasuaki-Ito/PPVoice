"""PowerPoint自動スピーチツール GUI (customtkinter)"""

import os
import re
import sys
import threading
import tkinter as tk
import winsound
from tkinter import colorchooser, filedialog, font as tkfont, messagebox

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import customtkinter as ctk

from pptx_reader import read_slides
from pptx_writer import embed_audio
from tts.voicevox import VoicevoxEngine
from version import __version__

ctk.set_appearance_mode("light")
ctk.set_default_color_theme(
    os.path.join(os.path.dirname(__file__), "theme_modern.json")
)


class _CancelledError(Exception):
    """生成処理の中断を伝える例外"""
    pass


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


class App(ctk.CTk):
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
        self._speaker_map: dict[str, int] = {}
        self._styles_by_speaker: dict[str, list[tuple[str, int]]] = {}
        self._running = False
        self._cancel_event = threading.Event()
        self._pending_speaker: str | None = None
        self._pending_style: str | None = None

        self._build_ui()
        self.input_var.trace_add("write", lambda *_: self._update_run_btn())
        self._update_run_btn()

    # ------------------------------------------------------------------
    # UI構築
    # ------------------------------------------------------------------

    def _build_ui(self):
        # 2カラムレイアウト: 左=設定, 右=生成・ログ
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=16, pady=16)
        container.grid_columnconfigure(0, weight=3, minsize=560)
        container.grid_columnconfigure(1, weight=2, minsize=280)
        container.grid_rowconfigure(0, weight=1)

        # 左カラム (スクロール可能な設定パネル)
        left = ctk.CTkScrollableFrame(container, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self._build_file_section(left)
        self._build_voice_section(left)
        self._build_subtitle_section(left)

        # 設定保存ボタン (左カラム最下部)
        ctk.CTkButton(
            left, text="設定保存", width=100, height=30, command=self._on_save_config,
        ).pack(anchor="e", padx=14, pady=(8, 4))

        # 右カラム (生成・ログ)
        right = ctk.CTkFrame(container, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self._build_action_section(right)

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

        self._section_header(sec, "ファイル設定")

        # 入力ファイル
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="入力ファイル (.pptx)", width=140, anchor="w").pack(side="left")
        self.input_var = ctk.StringVar()
        ctk.CTkEntry(row, textvariable=self.input_var).pack(side="left", fill="x", expand=True, padx=(4, 6))
        ctk.CTkButton(row, text="参照", width=60, command=self._browse_input).pack(side="left")

        # 出力ファイル
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="出力ファイル (.pptx)", width=140, anchor="w").pack(side="left")
        self.output_var = ctk.StringVar()
        ctk.CTkEntry(row, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=(4, 6))
        ctk.CTkButton(row, text="参照", width=60, command=self._browse_output).pack(side="left")

        # スライド範囲
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(3, 12))
        ctk.CTkLabel(row, text="スライド", width=100, anchor="w").pack(side="left")
        self.slide_range_var = ctk.StringVar(value="all")
        self._selected_slides: set[int] | None = None  # None = 全スライド
        ctk.CTkRadioButton(
            row, text="全部", variable=self.slide_range_var, value="all",
            command=self._on_slide_range_changed,
        ).pack(side="left", padx=(0, 16))
        ctk.CTkRadioButton(
            row, text="一部", variable=self.slide_range_var, value="select",
            command=self._on_slide_range_changed,
        ).pack(side="left", padx=(0, 8))
        self.slide_select_btn = ctk.CTkButton(
            row, text="選択...", width=60, command=self._open_slide_selector,
        )
        self.slide_select_label = ctk.CTkLabel(row, text="", anchor="w")
        # 初期状態では非表示
        self.slide_select_btn.pack_forget()
        self.slide_select_label.pack_forget()

    # --- 音声設定 ---
    def _build_voice_section(self, parent):
        sec = ctk.CTkFrame(parent)
        sec.pack(fill="x", pady=(0, 12))

        self._section_header(sec, "音声設定")

        # VOICEVOX URL
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="VOICEVOX URL", width=120, anchor="w").pack(side="left")
        self.url_var = ctk.StringVar(value="http://localhost:50021")
        ctk.CTkEntry(row, textvariable=self.url_var).pack(side="left", fill="x", expand=True, padx=(4, 6))
        ctk.CTkButton(row, text="話者取得", width=80, command=self._fetch_speakers).pack(side="left")

        # 話者選択
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="話者", width=120, anchor="w").pack(side="left")
        self.speaker_menu = ctk.CTkComboBox(
            row, values=["(話者取得を押してください)"],
            command=self._on_speaker_changed, state="readonly",
        )
        self.speaker_menu.pack(side="left", fill="x", expand=True, padx=(4, 0))

        # スタイル選択
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="スタイル", width=120, anchor="w").pack(side="left")
        self.style_speaker_menu = ctk.CTkComboBox(
            row, values=["---"], state="readonly",
        )
        self.style_speaker_menu.pack(side="left", fill="x", expand=True, padx=(4, 0))

        # テスト再生
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="テスト", width=120, anchor="w").pack(side="left")
        self.test_text_var = ctk.StringVar(value="音声のテストです。")
        ctk.CTkEntry(
            row, textvariable=self.test_text_var,
        ).pack(side="left", fill="x", expand=True, padx=(4, 6))
        self.test_play_btn = ctk.CTkButton(
            row, text="▶ 再生", width=80, command=self._on_test_play,
            state="disabled",
        )
        self.test_play_btn.pack(side="left")

        # VOICEVOX 利用規約リンク
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(0, 3))
        # 120px のスペーサーでラベル列と揃える
        ctk.CTkLabel(row, text="", width=120).pack(side="left")
        note = ctk.CTkLabel(
            row,
            text="※ キャラクターごとに利用規約があります → VOICEVOX公式サイト",
            text_color=None,
            font=ctk.CTkFont(size=11),
            cursor="hand2",
        )
        note.pack(side="left", padx=(4, 0))
        note.bind("<Button-1>", lambda e: __import__("webbrowser").open("https://voicevox.hiroshiba.jp/"))

        # 文間の間
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="文間の間 (秒)", width=120, anchor="w").pack(side="left")
        self.pause_var = ctk.DoubleVar(value=0.5)
        ctk.CTkSlider(row, from_=0.0, to=3.0, number_of_steps=30, variable=self.pause_var).pack(
            side="left", fill="x", expand=True, padx=(4, 8)
        )
        self.pause_label = ctk.CTkLabel(row, text="0.5", width=40)
        self.pause_label.pack(side="left")
        self.pause_var.trace_add("write", lambda *_: self.pause_label.configure(text=f"{self.pause_var.get():.1f}"))

        # 終了後の間
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(3, 12))
        ctk.CTkLabel(row, text="終了後の間 (秒)", width=120, anchor="w").pack(side="left")
        self.end_pause_var = ctk.DoubleVar(value=2.0)
        ctk.CTkSlider(row, from_=0.0, to=10.0, number_of_steps=100, variable=self.end_pause_var).pack(
            side="left", fill="x", expand=True, padx=(4, 8)
        )
        self.end_pause_label = ctk.CTkLabel(row, text="2.0", width=40)
        self.end_pause_label.pack(side="left")
        self.end_pause_var.trace_add(
            "write", lambda *_: self.end_pause_label.configure(text=f"{self.end_pause_var.get():.1f}")
        )

    # --- 字幕設定 ---
    def _build_subtitle_section(self, parent):
        sec = ctk.CTkFrame(parent)
        sec.pack(fill="x", pady=(0, 12))

        self._section_header(sec, "字幕設定")

        # 字幕有効
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        self.subtitle_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(row, text="字幕を表示する", variable=self.subtitle_var).pack(side="left")

        # スタイル
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="スタイル", width=120, anchor="w").pack(side="left")
        self.style_var = ctk.StringVar(value="outline")
        ctk.CTkRadioButton(
            row, text="縁取り", variable=self.style_var, value="outline",
            command=self._on_style_changed,
        ).pack(side="left", padx=(0, 16))
        ctk.CTkRadioButton(
            row, text="背景付き", variable=self.style_var, value="box",
            command=self._on_style_changed,
        ).pack(side="left")

        # フォントサイズ
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="フォントサイズ", width=120, anchor="w").pack(side="left")
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
        ctk.CTkLabel(row, text="フォント", width=120, anchor="w").pack(side="left")
        self._font_default_label = "<テーマのデフォルト>"
        self.subtitle_font_var = ctk.StringVar(value=self._font_default_label)
        ctk.CTkEntry(row, textvariable=self.subtitle_font_var).pack(side="left", fill="x", expand=True, padx=(4, 6))
        ctk.CTkButton(row, text="選択", width=60, command=self._open_font_picker).pack(side="left")

        # 下マージン
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="下マージン", width=120, anchor="w").pack(side="left")
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
        ctk.CTkLabel(row, text="文字色", width=120, anchor="w").pack(side="left")
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
            self.outline_row, text="輪郭", variable=self.use_outline_var, width=120,
        ).pack(side="left")
        self.outline_color_var = ctk.StringVar(value="#000000")
        self.outline_color_btn = ctk.CTkButton(
            self.outline_row, text="#000000", width=90, fg_color="#000000", text_color="#FFFFFF",
            command=lambda: self._pick_color(self.outline_color_var, self.outline_color_btn),
        )
        self.outline_color_btn.pack(side="left", padx=(4, 12))
        ctk.CTkLabel(self.outline_row, text="太さ", width=30, anchor="w").pack(side="left")
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
            self.glow_row, text="ぼかし", variable=self.use_glow_var, width=120,
        ).pack(side="left")
        self.glow_color_var = ctk.StringVar(value="#000000")
        self.glow_color_btn = ctk.CTkButton(
            self.glow_row, text="#000000", width=90, fg_color="#000000", text_color="#FFFFFF",
            command=lambda: self._pick_color(self.glow_color_var, self.glow_color_btn),
        )
        self.glow_color_btn.pack(side="left", padx=(4, 12))
        ctk.CTkLabel(self.glow_row, text="サイズ", width=40, anchor="w").pack(side="left")
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
        ctk.CTkLabel(self.bg_row, text="背景色", width=120, anchor="w").pack(side="left")
        self.bg_color_var = ctk.StringVar(value="#000000")
        self.bg_color_btn = ctk.CTkButton(
            self.bg_row, text="#000000", width=90, fg_color="#000000", text_color="#FFFFFF",
            command=lambda: self._pick_color(self.bg_color_var, self.bg_color_btn),
        )
        self.bg_color_btn.pack(side="left", padx=(4, 16))

        ctk.CTkLabel(self.bg_row, text="不透明度", width=60, anchor="w").pack(side="left")
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

    # --- 実行 / ログ ---
    def _build_action_section(self, parent):
        sec = ctk.CTkFrame(parent)
        sec.pack(fill="both", expand=True)

        self.run_btn = ctk.CTkButton(
            sec, text="生成開始", font=ctk.CTkFont(size=15, weight="bold"),
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
            self.input_var.set(path)
            if not self.output_var.get():
                base = os.path.splitext(path)[0]
                self.output_var.set(base + "_speech.pptx")
            # <config> タグの自動読み込み
            try:
                slides = read_slides(path)
                notes = [s.notes_text for s in slides if s.notes_text]
                config = self._parse_config_tags(notes)
                if config:
                    self._apply_config(config)
                    self._log(f"設定タグを読み込みました。\n")
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
            self._log("スライド選択にはまず入力ファイルを指定してください。\n")
            return
        try:
            slides = read_slides(input_path)
        except Exception as e:
            self._log(f"ファイル読み込み失敗: {e}\n")
            return
        total = len(slides)
        if total == 0:
            self._log("スライドが見つかりませんでした。\n")
            return

        # ポップアップウィンドウ
        dialog = ctk.CTkToplevel(self)
        dialog.title("スライド選択")
        dialog.geometry("360x420")
        dialog.resizable(False, True)
        dialog.grab_set()

        ctk.CTkLabel(
            dialog, text=f"作成するスライドを選択 ({total}枚)",
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

        ctk.CTkButton(btn_row, text="全選択", width=70, command=select_all).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_row, text="全解除", width=70, command=deselect_all).pack(side="left")

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
            label = f"スライド {num}"
            if preview:
                label += f" - {preview}"
            ctk.CTkCheckBox(scroll, text=label, variable=var).pack(anchor="w", pady=1)

        # OKボタン
        def on_ok():
            selected = {i + 1 for i, v in enumerate(check_vars) if v.get()}
            if not selected:
                messagebox.showwarning("選択なし", "少なくとも1枚選択してください。", parent=dialog)
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

    def _fetch_speakers(self):
        self._log_clear()
        url = self.url_var.get().strip().rstrip("/")
        try:
            engine = VoicevoxEngine(base_url=url)
            speakers = engine.list_speakers()
        except Exception as e:
            self._log(f"話者取得失敗: {e}\nVOICEVOXエンジンが起動しているか確認してください。\n")
            return

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
            self._log(f"{len(speakers)} 話者 ({total_styles} スタイル) を取得しました。\n")
            # pending speaker の適用
            self._apply_pending_speaker()
        else:
            self._log("話者が見つかりませんでした。\n")
        self._update_run_btn()

    def _on_speaker_changed(self, speaker_name: str):
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
    # テスト再生
    # ------------------------------------------------------------------

    def _on_test_play(self):
        btn_text = self.test_play_btn.cget("text")
        if btn_text == "■ 停止":
            winsound.PlaySound(None, winsound.SND_PURGE)
            self._test_play_reset()
            return

        text = self.test_text_var.get().strip()
        if not text:
            return
        if not self._speaker_map:
            self._log("先に話者を取得してください。\n")
            return

        style_label = self.style_speaker_menu.get()
        speaker_id = self._speaker_map.get(style_label, 1)
        url = self.url_var.get().strip()

        self.test_play_btn.configure(text="合成中...", state="disabled")
        thread = threading.Thread(
            target=self._test_play_worker,
            args=(text, speaker_id, url),
            daemon=True,
        )
        thread.start()

    def _test_play_worker(self, text, speaker_id, url):
        try:
            engine = VoicevoxEngine(speaker_id=speaker_id, base_url=url)
            wav = engine.synthesize(text)
            self.after(0, lambda: self.test_play_btn.configure(text="■ 停止", state="normal"))
            # SND_MEMORY は同期再生 (再生完了 or SND_PURGE で停止するまでブロック)
            winsound.PlaySound(wav, winsound.SND_MEMORY)
        except Exception as e:
            self.after(0, lambda: self._log(f"テスト再生エラー: {e}\n"))
        finally:
            self.after(0, self._test_play_reset)

    def _test_play_reset(self):
        state = "normal" if self._speaker_map else "disabled"
        self.test_play_btn.configure(text="▶ 再生", state=state)

    # ------------------------------------------------------------------
    # <config> タグ
    # ------------------------------------------------------------------

    _CONFIG_RE = re.compile(r"<config\s([^>]*)>", re.IGNORECASE)
    _KV_RE = re.compile(r'([\w]+)=(?:"([^"]*)"|(\S+))')

    def _parse_config_tags(self, notes_list: list[str]) -> dict:
        """複数のノートテキストから <config ...> タグを解析し設定 dict を返す。"""
        config: dict[str, str] = {}
        for notes in notes_list:
            for m in self._CONFIG_RE.finditer(notes):
                for kv in self._KV_RE.finditer(m.group(1)):
                    key = kv.group(1)
                    val = kv.group(2) if kv.group(2) is not None else kv.group(3)
                    config[key] = val
        return config

    def _apply_config(self, config: dict):
        """解析済み config dict を GUI ウィジェットに適用する。"""
        # 話者は pending に保存 (一覧取得後に適用)
        if "speaker" in config:
            self._pending_speaker = config["speaker"]
        if "style" in config:
            self._pending_style = config["style"]
        # すでに話者一覧がある場合は即適用
        if self._styles_by_speaker:
            self._apply_pending_speaker()

        # --- 音声設定 ---
        if "pause" in config:
            self.pause_var.set(float(config["pause"]))
        if "end_pause" in config:
            self.end_pause_var.set(float(config["end_pause"]))

        # --- 字幕設定 ---
        if "subtitle" in config:
            self.subtitle_var.set(config["subtitle"].lower() in ("on", "true", "1"))
        if "subtitle_style" in config:
            self.style_var.set(config["subtitle_style"])
            self._on_style_changed()
        if "fontsize" in config:
            self.fontsize_var.set(int(config["fontsize"]))
        if "font" in config:
            val = config["font"]
            self.subtitle_font_var.set(val if val else self._font_default_label)
        if "bottom" in config:
            self.bottom_var.set(float(config["bottom"]))
        if "font_color" in config:
            self._set_color(config["font_color"], self.font_color_var, self.font_color_btn)
        if "outline" in config:
            self.use_outline_var.set(config["outline"].lower() in ("on", "true", "1"))
        if "outline_color" in config:
            self._set_color(config["outline_color"], self.outline_color_var, self.outline_color_btn)
        if "outline_width" in config:
            self.outline_width_var.set(float(config["outline_width"]))
        if "glow" in config:
            self.use_glow_var.set(config["glow"].lower() in ("on", "true", "1"))
        if "glow_color" in config:
            self._set_color(config["glow_color"], self.glow_color_var, self.glow_color_btn)
        if "glow_size" in config:
            self.glow_size_var.set(float(config["glow_size"]))
        if "bg_color" in config:
            self._set_color(config["bg_color"], self.bg_color_var, self.bg_color_btn)
        if "bg_alpha" in config:
            self.bg_alpha_var.set(int(config["bg_alpha"]))

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

    def _apply_pending_speaker(self):
        """pending の話者・スタイル名を一覧から探して選択する。"""
        if self._pending_speaker:
            for display_name in (self.speaker_menu.cget("values") or []):
                name = display_name.rsplit(" (", 1)[0]
                if name == self._pending_speaker:
                    self.speaker_menu.set(display_name)
                    self._on_speaker_changed(display_name)
                    break
            self._pending_speaker = None

        if self._pending_style:
            for label in (self.style_speaker_menu.cget("values") or []):
                # "スタイル名 (ID=X)" から先頭のスタイル名を取得
                style_name = label.rsplit(" (ID=", 1)[0]
                if style_name == self._pending_style:
                    self.style_speaker_menu.set(label)
                    break
            self._pending_style = None

    def _generate_config_tag(self) -> str:
        """現在の GUI 設定を <config ...> タグ文字列として生成する。"""
        parts: list[str] = []

        def _add(key, val):
            s = str(val)
            if " " in s or not s:
                parts.append(f'{key}="{s}"')
            else:
                parts.append(f"{key}={s}")

        # 話者
        speaker_display = self.speaker_menu.get()
        if speaker_display and " (" in speaker_display:
            _add("speaker", speaker_display.rsplit(" (", 1)[0])
        style_display = self.style_speaker_menu.get()
        if style_display and " (ID=" in style_display:
            _add("style", style_display.rsplit(" (ID=", 1)[0])

        # 音声
        _add("pause", f"{self.pause_var.get():.1f}")
        _add("end_pause", f"{self.end_pause_var.get():.1f}")

        # 字幕
        _add("subtitle", "on" if self.subtitle_var.get() else "off")
        _add("subtitle_style", self.style_var.get())
        _add("fontsize", self.fontsize_var.get())
        font_val = self.subtitle_font_var.get()
        _add("font", "" if font_val == self._font_default_label else font_val)
        _add("bottom", f"{self.bottom_var.get():.2f}")
        _add("font_color", self.font_color_var.get())
        _add("outline", "on" if self.use_outline_var.get() else "off")
        _add("outline_color", self.outline_color_var.get())
        _add("outline_width", f"{self.outline_width_var.get():.2f}")
        _add("glow", "on" if self.use_glow_var.get() else "off")
        _add("glow_color", self.glow_color_var.get())
        _add("glow_size", f"{self.glow_size_var.get():.1f}")
        _add("bg_color", self.bg_color_var.get())
        _add("bg_alpha", self.bg_alpha_var.get())

        return "<config " + " ".join(parts) + ">"

    def _on_save_config(self):
        """設定保存ポップアップを表示する。"""
        tag = self._generate_config_tag()

        dialog = ctk.CTkToplevel(self)
        dialog.title("設定保存")
        dialog.geometry("600x220")
        dialog.resizable(True, False)
        dialog.grab_set()

        ctk.CTkLabel(
            dialog,
            text="以下のタグを PPTX の任意のスライドのノート欄に貼り付けると、\nファイルを開いた際に設定が自動的に読み込まれます。",
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
            copy_btn.configure(text="コピーしました")
            dialog.after(1500, lambda: copy_btn.configure(text="コピー"))

        copy_btn = ctk.CTkButton(dialog, text="コピー", width=120, command=on_copy)
        copy_btn.pack(pady=(0, 16))

    def _open_font_picker(self):
        families = sorted(
            (f for f in set(tkfont.families()) if not f.startswith("@")),
            key=str.lower,
        )
        all_fonts = [self._font_default_label] + families

        dialog = ctk.CTkToplevel(self)
        dialog.title("フォント選択")
        dialog.geometry("400x500")
        dialog.resizable(True, True)
        dialog.grab_set()

        # 検索
        search_var = ctk.StringVar()
        ctk.CTkEntry(
            dialog, textvariable=search_var, placeholder_text="検索...",
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
        color = colorchooser.askcolor(color=var.get(), title="色を選択")
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
        if btn_text == "▶ 再生":
            self.test_play_btn.configure(state="normal" if speaker_ok else "disabled")

    def _on_run(self):
        if self._running:
            # 停止要求
            self._cancel_event.set()
            self.run_btn.configure(state="disabled", text="停止中...")
            return

        input_path = self.input_var.get().strip()
        if not input_path:
            self._log("入力ファイルを指定してください。\n")
            return
        if not os.path.exists(input_path):
            self._log(f"ファイルが見つかりません: {input_path}\n")
            return

        # 出力ファイルの上書き確認
        base_name = os.path.splitext(input_path)[0]
        output_path = self.output_var.get().strip()
        if not output_path:
            output_path = base_name + "_speech.pptx"
        existing = [output_path] if os.path.exists(output_path) else []
        if existing:
            names = "\n".join(os.path.basename(f) for f in existing)
            if not messagebox.askyesno("上書き確認", f"以下のファイルが既に存在します。上書きしますか?\n\n{names}"):
                return

        self._running = True
        self._cancel_event.clear()
        self._log_clear()
        self.run_btn.configure(text="停止", fg_color="#EF4444", hover_color="#DC2626")
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
            print("\n処理を中断しました。")
        except Exception as e:
            print(f"\nエラー: {e}")
        finally:
            sys.stdout = old_stdout
            self.after(0, self._on_done)

    def _on_done(self):
        self._running = False
        self._cancel_event.clear()
        self.progress.pack_forget()
        self.run_btn.configure(
            text="生成開始",
            fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"],
            hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"],
        )
        self._update_run_btn()

    def _do_generate(self):
        input_path = self.input_var.get().strip()
        base_name = os.path.splitext(input_path)[0]

        output_path = self.output_var.get().strip()
        if not output_path:
            output_path = base_name + "_speech.pptx"

        # 話者ID
        style_label = self.style_speaker_menu.get()
        speaker_id = self._speaker_map.get(style_label, 1)

        url = self.url_var.get().strip()
        pause_sec = self.pause_var.get()
        end_pause_sec = self.end_pause_var.get()
        use_subtitle = self.subtitle_var.get()
        sub_style = self.style_var.get()
        sub_size = self.fontsize_var.get()
        sub_bottom = self.bottom_var.get()
        _font_sel = self.subtitle_font_var.get().strip()
        sub_font_name = "" if _font_sel == self._font_default_label else _font_sel
        sub_font_color = self.font_color_var.get().lstrip("#")
        sub_use_outline = self.use_outline_var.get()
        sub_outline_color = self.outline_color_var.get().lstrip("#")
        sub_outline_width = self.outline_width_var.get()
        sub_use_glow = self.use_glow_var.get()
        sub_glow_color = self.glow_color_var.get().lstrip("#")
        sub_glow_size = self.glow_size_var.get()
        sub_bg_color = self.bg_color_var.get().lstrip("#")
        sub_bg_alpha = self.bg_alpha_var.get()

        # スライド読み込み
        print(f"PPTXを読み込んでいます: {input_path}")
        slides = read_slides(input_path)
        total_slides = len(slides)
        print(f"  {total_slides} スライドを検出")

        # スライドフィルタ
        selected = self._selected_slides
        if selected is not None:
            slides = [s for s in slides if (s.index + 1) in selected]
            print(f"  {len(slides)} スライドを選択中")

        notes_count = sum(1 for s in slides if s.notes_text)
        if notes_count == 0:
            print("ノートが含まれるスライドがありません。終了します。")
            return
        print(f"  {notes_count} スライドにノートあり")

        # 音声合成
        need_timings = use_subtitle
        print(f"\n音声を合成しています (speaker={speaker_id}, pause={pause_sec}s)...")
        engine = VoicevoxEngine(speaker_id=speaker_id, base_url=url, pause_sec=pause_sec)

        slide_audio = []
        slide_timings = {}
        processed = 0

        for info in slides:
            if self._cancel_event.is_set():
                raise _CancelledError()

            slide_num = info.index + 1
            if not info.notes_text:
                print(f"  [{slide_num}/{total_slides}] スライド {slide_num}: (ノートなし - スキップ)")
                slide_audio.append((info.index, b""))
            else:
                print(f"  [{slide_num}/{total_slides}] スライド {slide_num}:")

                def on_chunk(i, total, text, _sn=slide_num):
                    if self._cancel_event.is_set():
                        raise _CancelledError()
                    print(f"    ({i + 1}/{total}) {text}")

                if need_timings:
                    wav, timings = engine.synthesize_with_timings(info.notes_text, on_chunk=on_chunk)
                    slide_timings[info.index] = timings
                else:
                    wav = engine.synthesize(info.notes_text, on_chunk=on_chunk)
                slide_audio.append((info.index, wav))

            processed += 1
            self.after(0, self.progress.set, processed / total_slides)

        # PPTX出力
        print(f"\n音声付きPPTXを生成しています...")
        embed_audio(
            input_path,
            slide_audio,
            output_path,
            end_pause_ms=int(end_pause_sec * 1000),
            slide_timings=slide_timings if need_timings else None,
            subtitle_font_size=sub_size,
            subtitle_font_name=sub_font_name,
            subtitle_bottom_pct=sub_bottom,
            subtitle_style=sub_style,
            subtitle_font_color=sub_font_color,
            subtitle_use_outline=sub_use_outline,
            subtitle_outline_color=sub_outline_color,
            subtitle_outline_width=sub_outline_width,
            subtitle_use_glow=sub_use_glow,
            subtitle_glow_color=sub_glow_color,
            subtitle_glow_size=sub_glow_size,
            subtitle_bg_color=sub_bg_color,
            subtitle_bg_alpha=sub_bg_alpha,
        )

        self.after(0, self.progress.set, 1.0)
        print(f"\n完了! → {os.path.basename(output_path)}")
        print("\n--- 動画 (MP4) にするには ---")
        print("1. 生成されたPPTXをPowerPointで開く")
        print("2. ファイル → エクスポート → ビデオの作成")
        print("3. 品質を選択して「ビデオの作成」をクリック")


if __name__ == "__main__":
    app = App()
    app.mainloop()
