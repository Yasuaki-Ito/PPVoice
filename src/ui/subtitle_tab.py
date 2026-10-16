"""字幕タブ: 表示・文字・装飾・位置・句読点"""

import tkinter as tk
from tkinter import font as tkfont

import customtkinter as ctk

import i18n
from i18n import t
from settings import Settings
from ui import widgets as w

# 句読点の置換先: (内部の値, 表示用の文言キー)。内部の値は <config> の kuten/touten に保存する値で、
# 既存の PPTX との互換のため日本語の表記のまま。表示用の文言キーが None なら値をそのまま表示する
TOUTEN_CHOICES = [("そのまま", "punct_unchanged"), ("、", None), (",(半角)", "punct_comma_half"),
                  ("，(全角)", "punct_comma_full"), ("(半角空白)", "punct_space_half"), ("(全角空白)", "punct_space_full")]
KUTEN_CHOICES = [("そのまま", "punct_unchanged"), ("。", None), (".(半角)", "punct_period_half"),
                 ("．(全角)", "punct_period_full"), ("(半角空白)", "punct_space_half"), ("(全角空白)", "punct_space_full")]


def punct_label(choices, value: str) -> str:
    """句読点の置換先の内部の値 (または日英いずれかの表示名) → 現在の言語の表示名"""
    for v, key in choices:
        names = (v,) if key is None else (v, *i18n.STRINGS[key])
        if value in names:
            return v if key is None else t(key)
    return value  # 一覧にない 1 文字の指定などはそのまま


def punct_value(choices, text: str) -> str:
    """表示名 → 句読点の置換先の内部の値"""
    for v, key in choices:
        if text == v or (key is not None and text in i18n.STRINGS[key]):
            return v
    return text


class SubtitleTab(ctk.CTkScrollableFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self._build()

    def _build(self):
        r = w.row(self, pady=(8, 2))
        self.subtitle_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(r, text=t("show_subtitles"), variable=self.subtitle_var).pack(side="left", padx=(0, 24))
        self.name_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(r, text=t("show_speaker_name"), variable=self.name_var).pack(side="left")

        # 文字
        w.heading(self, t("sub_text"))
        r = w.row(self)
        w.label(r, t("font"))
        self.font_default_label = t("font_default")
        self.font_var = ctk.StringVar(value=self.font_default_label)
        ctk.CTkEntry(r, textvariable=self.font_var).pack(side="left", fill="x", expand=True, padx=(4, 6))
        ctk.CTkButton(r, text=t("select"), width=60, command=self.open_font_picker).pack(side="left")
        self.fontsize_var = ctk.IntVar(value=18)
        w.SliderRow(self, t("font_size"), self.fontsize_var, 10, 48, 38, "{:.0f}").pack(
            fill="x", padx=12, pady=w.ROW_PAD)
        r = w.row(self)
        w.label(r, t("font_color"))
        self.font_color_var = ctk.StringVar(value="#FFFFFF")
        w.ColorButton(r, self.font_color_var, command=self.app.on_font_color_changed).pack(side="left", padx=(4, 12))
        w.hint(r, t("role_color_hint")).pack(side="left")
        r = w.row(self)
        w.label(r, t("decoration"))
        self.bold_var = ctk.BooleanVar(value=False)
        self.italic_var = ctk.BooleanVar(value=False)
        self.underline_var = ctk.BooleanVar(value=False)
        self.math_bold_var = ctk.BooleanVar(value=True)
        for text, var in ((t("bold"), self.bold_var), (t("italic"), self.italic_var),
                          (t("underline"), self.underline_var), (t("math_bold"), self.math_bold_var)):
            ctk.CTkCheckBox(r, text=text, variable=var, width=60).pack(side="left", padx=(4, 8))

        # 装飾
        w.heading(self, t("sub_decor"))
        r = w.row(self)
        w.label(r, t("sub_style"))
        self.style_var = ctk.StringVar(value="outline")
        for text, value in ((t("style_outline"), "outline"), (t("style_box"), "box")):
            ctk.CTkRadioButton(r, text=text, variable=self.style_var, value=value,
                               command=self._on_style_changed).pack(side="left", padx=(4, 16))
        self.decor = ctk.CTkFrame(self, fg_color="transparent")
        self.decor.pack(fill="x")
        # 縁取り: 輪郭・ぼかし
        self.outline_row = r = ctk.CTkFrame(self.decor, fg_color="transparent")
        self.outline_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(r, text=t("outline"), variable=self.outline_var, width=w.LABEL_WIDTH).pack(side="left")
        self.outline_color_var = ctk.StringVar(value="#000000")
        w.ColorButton(r, self.outline_color_var).pack(side="left", padx=(4, 12))
        self.outline_width_var = ctk.DoubleVar(value=0.75)
        w.SliderRow(r, t("outline_width"), self.outline_width_var, 0.25, 6.0, 23, "{:.2f}", label_width=40,
                    slider_width=140).pack(side="left")
        self.glow_row = r = ctk.CTkFrame(self.decor, fg_color="transparent")
        self.glow_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(r, text=t("glow"), variable=self.glow_var, width=w.LABEL_WIDTH).pack(side="left")
        self.glow_color_var = ctk.StringVar(value="#000000")
        w.ColorButton(r, self.glow_color_var).pack(side="left", padx=(4, 12))
        self.glow_size_var = ctk.DoubleVar(value=11.0)
        w.SliderRow(r, t("glow_size"), self.glow_size_var, 1.0, 30.0, 29, label_width=40,
                    slider_width=140).pack(side="left")
        # 背景付き: 背景色・不透明度
        self.bg_row = r = ctk.CTkFrame(self.decor, fg_color="transparent")
        w.label(r, t("bg_color"))
        self.bg_color_var = ctk.StringVar(value="#000000")
        w.ColorButton(r, self.bg_color_var).pack(side="left", padx=(4, 12))
        self.bg_alpha_var = ctk.IntVar(value=60)
        w.SliderRow(r, t("bg_alpha"), self.bg_alpha_var, 0, 100, 100, "{:.0f}%", label_width=60,
                    slider_width=140).pack(side="left")

        # 位置
        w.heading(self, t("sub_position"))
        self.bottom_var = ctk.DoubleVar(value=0.05)
        w.SliderRow(self, t("bottom_margin"), self.bottom_var, 0.0, 0.3, 30, "{:.2f}").pack(
            fill="x", padx=12, pady=w.ROW_PAD)

        # 句読点
        w.heading(self, t("sub_punct"))
        r = w.row(self, pady=(2, 10))
        w.label(r, t("touten"))
        self.touten_var = ctk.StringVar(value=punct_label(TOUTEN_CHOICES, "そのまま"))
        ctk.CTkComboBox(r, values=[punct_label(TOUTEN_CHOICES, v) for v, _ in TOUTEN_CHOICES],
                        variable=self.touten_var, state="readonly", width=150).pack(side="left", padx=(4, 16))
        ctk.CTkLabel(r, text=t("kuten")).pack(side="left")
        self.kuten_var = ctk.StringVar(value=punct_label(KUTEN_CHOICES, "そのまま"))
        ctk.CTkComboBox(r, values=[punct_label(KUTEN_CHOICES, v) for v, _ in KUTEN_CHOICES],
                        variable=self.kuten_var, state="readonly", width=150).pack(side="left", padx=(4, 0))

        self._on_style_changed()

    def _on_style_changed(self):
        for r in (self.outline_row, self.glow_row, self.bg_row):
            r.pack_forget()
        rows = (self.outline_row, self.glow_row) if self.style_var.get() == "outline" else (self.bg_row,)
        for r in rows:
            r.pack(fill="x", padx=12, pady=w.ROW_PAD)

    # ------------------------------------------------------------------
    # 設定の読み書き
    # ------------------------------------------------------------------

    def load(self, s: Settings):
        self.subtitle_var.set(s.subtitle)
        self.name_var.set(s.subtitle_name)
        self.font_var.set(s.font or self.font_default_label)
        self.fontsize_var.set(s.fontsize)
        self.font_color_var.set(s.font_color)
        self.bold_var.set(s.bold)
        self.italic_var.set(s.italic)
        self.underline_var.set(s.underline)
        self.math_bold_var.set(s.math_bold)
        self.style_var.set(s.subtitle_style)
        self._on_style_changed()
        self.outline_var.set(s.outline)
        self.outline_color_var.set(s.outline_color)
        self.outline_width_var.set(s.outline_width)
        self.glow_var.set(s.glow)
        self.glow_color_var.set(s.glow_color)
        self.glow_size_var.set(s.glow_size)
        self.bg_color_var.set(s.bg_color)
        self.bg_alpha_var.set(s.bg_alpha)
        self.bottom_var.set(s.bottom)
        self.touten_var.set(punct_label(TOUTEN_CHOICES, s.touten))
        self.kuten_var.set(punct_label(KUTEN_CHOICES, s.kuten))

    def save(self, s: Settings):
        s.subtitle = self.subtitle_var.get()
        s.subtitle_name = self.name_var.get()
        font = self.font_var.get().strip()
        s.font = "" if font == self.font_default_label else font
        s.fontsize = int(self.fontsize_var.get())
        s.font_color = self.font_color_var.get()
        s.bold = self.bold_var.get()
        s.italic = self.italic_var.get()
        s.underline = self.underline_var.get()
        s.math_bold = self.math_bold_var.get()
        s.subtitle_style = self.style_var.get()
        s.outline = self.outline_var.get()
        s.outline_color = self.outline_color_var.get()
        s.outline_width = self.outline_width_var.get()
        s.glow = self.glow_var.get()
        s.glow_color = self.glow_color_var.get()
        s.glow_size = self.glow_size_var.get()
        s.bg_color = self.bg_color_var.get()
        s.bg_alpha = int(self.bg_alpha_var.get())
        s.bottom = self.bottom_var.get()
        s.touten = punct_value(TOUTEN_CHOICES, self.touten_var.get())
        s.kuten = punct_value(KUTEN_CHOICES, self.kuten_var.get())

    # ------------------------------------------------------------------
    # フォントの選択 (選んだフォントのプレビュー付き)
    # ------------------------------------------------------------------

    def open_font_picker(self):
        families = sorted((f for f in set(tkfont.families()) if not f.startswith("@")), key=str.lower)
        all_fonts = [self.font_default_label] + families

        dialog = ctk.CTkToplevel(self)
        dialog.title(t("font_picker_title"))
        dialog.geometry("420x560")
        dialog.grab_set()
        search_var = ctk.StringVar()
        ctk.CTkEntry(dialog, textvariable=search_var, placeholder_text=t("search")).pack(
            fill="x", padx=10, pady=(10, 6))
        list_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        listbox = tk.Listbox(list_frame, font=("", 11), activestyle="dotbox", exportselection=False)
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True)
        # プレビュー (tk はフォント名をそのまま使える)
        preview = tk.Label(dialog, text=t("font_preview"), font=("", 20), height=2, bg="#333333", fg="#FFFFFF")
        preview.pack(fill="x", padx=10, pady=(0, 6))

        def show_preview(_e=None):
            sel = listbox.curselection()
            name = listbox.get(sel[0]) if sel else ""
            preview.configure(font=("" if not name or name == self.font_default_label else name, 20))

        def populate(query=""):
            listbox.delete(0, "end")
            q = query.lower()
            for f in all_fonts:
                if q in f.lower():
                    listbox.insert("end", f)
            items = list(listbox.get(0, "end"))
            current = self.font_var.get()
            if current in items:
                idx = items.index(current)
                listbox.selection_set(idx)
                listbox.see(idx)
            show_preview()

        populate()
        search_var.trace_add("write", lambda *_: populate(search_var.get()))
        listbox.bind("<<ListboxSelect>>", show_preview)

        def on_ok():
            sel = listbox.curselection()
            if sel:
                self.font_var.set(listbox.get(sel[0]))
            dialog.destroy()

        listbox.bind("<Double-1>", lambda e: on_ok())
        ctk.CTkButton(dialog, text="OK", width=120, command=on_ok).pack(pady=(0, 10))
