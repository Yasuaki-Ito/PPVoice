"""PowerPoint自動スピーチツール GUI (customtkinter)"""

import os
import sys
import threading
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox

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
        self.geometry("700x820")
        self.minsize(600, 700)

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

        self._build_ui()
        self.input_var.trace_add("write", lambda *_: self._update_run_btn())
        self._update_run_btn()

    # ------------------------------------------------------------------
    # UI構築
    # ------------------------------------------------------------------

    def _build_ui(self):
        wrapper = ctk.CTkScrollableFrame(self, fg_color="transparent")
        wrapper.pack(fill="both", expand=True, padx=16, pady=16)

        self._build_file_section(wrapper)
        self._build_voice_section(wrapper)
        self._build_subtitle_section(wrapper)
        self._build_action_section(wrapper)

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
        ctk.CTkEntry(row, textvariable=self.input_var, width=380).pack(side="left", padx=(4, 6))
        ctk.CTkButton(row, text="参照", width=60, command=self._browse_input).pack(side="left")

        # 出力ファイル
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="出力ファイル (.pptx)", width=140, anchor="w").pack(side="left")
        self.output_var = ctk.StringVar()
        ctk.CTkEntry(row, textvariable=self.output_var, width=380).pack(side="left", padx=(4, 6))
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
        ctk.CTkEntry(row, textvariable=self.url_var, width=280).pack(side="left", padx=(4, 6))
        ctk.CTkButton(row, text="話者取得", width=80, command=self._fetch_speakers).pack(side="left")

        # 話者選択
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="話者", width=120, anchor="w").pack(side="left")
        self.speaker_menu = ctk.CTkComboBox(
            row, values=["(話者取得を押してください)"], width=380,
            command=self._on_speaker_changed, state="readonly",
        )
        self.speaker_menu.pack(side="left", padx=(4, 0))

        # スタイル選択
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="スタイル", width=120, anchor="w").pack(side="left")
        self.style_speaker_menu = ctk.CTkComboBox(
            row, values=["---"], width=380, state="readonly",
        )
        self.style_speaker_menu.pack(side="left", padx=(4, 0))

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
        ctk.CTkSlider(row, from_=0.0, to=3.0, number_of_steps=30, variable=self.pause_var, width=300).pack(
            side="left", padx=(4, 8)
        )
        self.pause_label = ctk.CTkLabel(row, text="0.5", width=40)
        self.pause_label.pack(side="left")
        self.pause_var.trace_add("write", lambda *_: self.pause_label.configure(text=f"{self.pause_var.get():.1f}"))

        # 終了後の間
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(3, 12))
        ctk.CTkLabel(row, text="終了後の間 (秒)", width=120, anchor="w").pack(side="left")
        self.end_pause_var = ctk.DoubleVar(value=2.0)
        ctk.CTkSlider(row, from_=0.0, to=10.0, number_of_steps=100, variable=self.end_pause_var, width=300).pack(
            side="left", padx=(4, 8)
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
        ctk.CTkSlider(row, from_=10, to=48, number_of_steps=38, variable=self.fontsize_var, width=300).pack(
            side="left", padx=(4, 8)
        )
        self.fontsize_label = ctk.CTkLabel(row, text="18", width=40)
        self.fontsize_label.pack(side="left")
        self.fontsize_var.trace_add(
            "write", lambda *_: self.fontsize_label.configure(text=str(self.fontsize_var.get()))
        )

        # 下マージン
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="下マージン", width=120, anchor="w").pack(side="left")
        self.bottom_var = ctk.DoubleVar(value=0.05)
        ctk.CTkSlider(row, from_=0.0, to=0.3, number_of_steps=30, variable=self.bottom_var, width=300).pack(
            side="left", padx=(4, 8)
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

        # 縁取り色 (outline 時のみ表示)
        self.glow_row = ctk.CTkFrame(sec, fg_color="transparent")
        ctk.CTkLabel(self.glow_row, text="縁取り色", width=120, anchor="w").pack(side="left")
        self.glow_color_var = ctk.StringVar(value="#000000")
        self.glow_color_btn = ctk.CTkButton(
            self.glow_row, text="#000000", width=90, fg_color="#000000", text_color="#FFFFFF",
            command=lambda: self._pick_color(self.glow_color_var, self.glow_color_btn),
        )
        self.glow_color_btn.pack(side="left", padx=(4, 0))

        # 背景色・透過度 (box 時のみ表示)
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
        sec.pack(fill="x", pady=(0, 0))

        self.run_btn = ctk.CTkButton(
            sec, text="生成開始", font=ctk.CTkFont(size=15, weight="bold"),
            height=44, corner_radius=12, command=self._on_run,
        )
        self.run_btn.pack(fill="x", padx=14, pady=(14, 8))

        self.progress = ctk.CTkProgressBar(sec, height=6, corner_radius=3)
        self.progress.pack(fill="x", padx=14, pady=(0, 8))
        self.progress.set(0)

        self.log_box = ctk.CTkTextbox(sec, height=200, state="disabled", font=ctk.CTkFont(size=12))
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

    def _on_style_changed(self):
        if self.style_var.get() == "outline":
            self.bg_row.pack_forget()
            self.glow_row.pack(fill="x", padx=14, pady=3)
        else:
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
        """入力ファイル・話者選択の状態に応じて生成ボタンの有効/無効を切り替える。"""
        if self._running:
            return
        input_ok = bool(self.input_var.get().strip())
        speaker_ok = bool(self._speaker_map)
        if input_ok and speaker_ok:
            self.run_btn.configure(state="normal")
        else:
            self.run_btn.configure(state="disabled")

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
        sub_font_color = self.font_color_var.get().lstrip("#")
        sub_glow_color = self.glow_color_var.get().lstrip("#")
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
            subtitle_bottom_pct=sub_bottom,
            subtitle_style=sub_style,
            subtitle_font_color=sub_font_color,
            subtitle_glow_color=sub_glow_color,
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
