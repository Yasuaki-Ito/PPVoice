"""GUI の共通部品 (ラベル + スライダー + 値の行、色ボタン、話者の2段の選択など)"""

from tkinter import colorchooser

import customtkinter as ctk

from i18n import t
from tts.openai_compat import KOKORO_LANG_CODES, group_voices, voice_short_name

LABEL_WIDTH = 110
ROW_PAD = 2
HINT_COLOR = ("gray45", "gray60")


def row(parent, pady=ROW_PAD, padx=12) -> ctk.CTkFrame:
    """横に並べる行"""
    r = ctk.CTkFrame(parent, fg_color="transparent")
    r.pack(fill="x", padx=padx, pady=pady)
    return r


def label(parent, text: str, width: int = LABEL_WIDTH, **kw) -> ctk.CTkLabel:
    lb = ctk.CTkLabel(parent, text=text, width=width, anchor="w", **kw)
    lb.pack(side="left")
    return lb


def heading(parent, text: str, pady=(10, 2)) -> ctk.CTkLabel:
    """まとまりの小見出し"""
    lb = ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=13, weight="bold"),
                      text_color=("#4C566A", "#B0B8C8"), anchor="w")
    lb.pack(fill="x", padx=12, pady=pady)
    return lb


def hint(parent, text: str, **kw) -> ctk.CTkLabel:
    lb = ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=11), text_color=HINT_COLOR, anchor="w", **kw)
    return lb


class SliderRow(ctk.CTkFrame):
    """ラベル + スライダー + 値の表示"""

    def __init__(self, parent, text: str, var, from_: float, to: float, steps: int, fmt: str = "{:.1f}",
                 label_width: int = LABEL_WIDTH, slider_width: int | None = None, command=None):
        super().__init__(parent, fg_color="transparent")
        self.var = var
        self.label = label(self, text, width=label_width)
        kw = {"width": slider_width} if slider_width else {}
        self.slider = ctk.CTkSlider(self, from_=from_, to=to, number_of_steps=steps, variable=var,
                                    command=(lambda _v: command()) if command else None, **kw)
        self.slider.pack(side="left", fill="none" if slider_width else "x", expand=not slider_width,
                         padx=(4, 8))
        self.value = ctk.CTkLabel(self, text=fmt.format(var.get()), width=40)
        self.value.pack(side="left")
        var.trace_add("write", lambda *_: self._show(fmt))

    def _show(self, fmt):
        try:
            self.value.configure(text=fmt.format(self.var.get()))
        except Exception:
            pass


def text_color_for(hex_color: str) -> str:
    """背景色の上で読める文字色 (白または黒)"""
    try:
        r, g, b = int(hex_color[1:3], 16), int(hex_color[3:5], 16), int(hex_color[5:7], 16)
    except (ValueError, IndexError):
        return "#000000"
    return "#000000" if (r * 0.299 + g * 0.587 + b * 0.114) > 128 else "#FFFFFF"


class ColorButton(ctk.CTkButton):
    """色を表示し、押すと色を選ぶボタン。allow_empty なら空 (= 既定の色) も選べる (右クリックで空に戻す)。"""

    def __init__(self, parent, var: ctk.StringVar, width: int = 90, empty_text: str = "", command=None):
        super().__init__(parent, width=width, command=self._pick, border_width=1,
                         border_color=("gray70", "gray40"))
        self.var = var
        self.empty_text = empty_text
        self._command = command
        if empty_text:
            self.bind("<Button-3>", lambda e: self.set(""))
        var.trace_add("write", lambda *_: self._show())
        self._show()

    def set(self, value: str):
        if value and not value.startswith("#"):
            value = "#" + value
        self.var.set(value.upper())
        if self._command:
            self._command()

    def _show(self):
        value = self.var.get()
        if not value:
            self.configure(text=self.empty_text, fg_color=("gray85", "gray30"), text_color=("gray20", "gray85"))
            return
        try:
            self.configure(text=value, fg_color=value, text_color=text_color_for(value))
        except Exception:
            pass

    def _pick(self):
        color = colorchooser.askcolor(color=self.var.get() or "#FFFFFF", title=t("color_picker_title"))
        if color[1]:
            self.set(color[1])


class VoicePicker:
    """話者の2段の選択 (VOICEVOX: 話者 → スタイル / Kokoro: 言語・性別 → 声 / その他の OpenAI 互換: 声だけ)

    話者は名前で扱う。一覧を取得する前は、選びたい名前を「(未取得) ...」と表示しておき、
    取得したらその名前の項目を選ぶ。
    """

    def __init__(self, parent, label_width: int = LABEL_WIDTH, on_change=None):
        self.engine = "voicevox"
        self._on_change = on_change
        self._styles_by_speaker: dict[str, list[tuple[str, int | str]]] = {}
        self._value_by_label: dict[str, int | str] = {}
        self._fetched = False
        self._grouped = False
        self._pending = ("", "")

        self.speaker_row = row(parent)
        self.speaker_label = label(self.speaker_row, t("speaker"), width=label_width)
        self.speaker_menu = ctk.CTkComboBox(self.speaker_row, values=[t("press_fetch")], state="readonly",
                                            command=self._on_speaker_selected)
        self.speaker_menu.pack(side="left", fill="x", expand=True, padx=(4, 0))
        self.style_row = row(parent)
        self.style_label = label(self.style_row, t("style"), width=label_width)
        self.style_menu = ctk.CTkComboBox(self.style_row, values=["---"], state="readonly",
                                          command=lambda _v: self._changed())
        self.style_menu.pack(side="left", fill="x", expand=True, padx=(4, 0))
        self._show_pending()

    # --- 一覧 ---
    @property
    def fetched(self) -> bool:
        return self._fetched

    def clear(self, engine: str):
        """エンジンを切り替えたとき (一覧は取り直し)"""
        self.engine = engine
        self._fetched = False
        self._grouped = False
        self._styles_by_speaker = {}
        self._value_by_label = {}
        self._apply_labels()
        self._show_pending()

    def set_voicevox(self, speakers: list[dict]):
        self.engine = "voicevox"
        self._grouped = False
        self._styles_by_speaker = {}
        for sp in speakers:
            styles = [(f"{st['name']} (ID={st['id']})", st["id"]) for st in sp.get("styles", [])]
            if styles:
                self._styles_by_speaker[sp["name"]] = styles
        self._fetched = bool(self._styles_by_speaker)
        self._apply_labels()
        self.speaker_menu.configure(values=[f"{n} ({len(s)})" for n, s in self._styles_by_speaker.items()])
        self.select(*self._pending)

    def set_openai(self, voices: list[str]):
        self.engine = "openai"
        groups = group_voices(voices)
        self._grouped = groups is not None
        self._styles_by_speaker = {}
        if groups:
            for key, ids in groups:
                self._styles_by_speaker[self.group_label(key)] = [(f"{voice_short_name(v)} ({v})", v) for v in ids]
            self.speaker_menu.configure(values=[f"{n} ({len(s)})" for n, s in self._styles_by_speaker.items()])
        else:
            self._value_by_label = {v: v for v in voices}
            self.speaker_menu.configure(values=list(voices) or ["---"])
        self._fetched = bool(voices)
        self._apply_labels()
        self.select(*self._pending)

    @staticmethod
    def group_label(key: tuple[str, str] | None) -> str:
        """Kokoro の声のグループ (言語の記号, 性別の記号) の表示名"""
        if key is None:
            return t("voice_group_other")
        lang, gender = key
        lang_name = t(f"kokoro_lang_{lang}") if lang in KOKORO_LANG_CODES else t("lang_unknown", code=lang)
        return t("voice_group_label", lang=lang_name, gender=t(f"gender_{gender}"))

    def _apply_labels(self):
        voicevox = self.engine == "voicevox"
        self.speaker_label.configure(
            text=t("speaker") if voicevox else (t("voice_group") if self._grouped else t("voice")))
        self.style_label.configure(text=t("style") if voicevox else t("voice"))
        if voicevox or self._grouped:
            self.style_row.pack(fill="x", padx=12, pady=ROW_PAD, after=self.speaker_row)
        else:
            self.style_row.pack_forget()

    def _show_pending(self):
        speaker, style = self._pending
        text = t("not_fetched", name=" / ".join(x for x in (speaker, style) if x)) if speaker else t("press_fetch")
        self.speaker_menu.configure(values=[text])
        self.speaker_menu.set(text)
        self.style_menu.configure(values=["---"])
        self.style_menu.set("---")

    # --- 選択 ---
    def _on_speaker_selected(self, display: str):
        self._fill_styles(display)
        self._changed()

    def _fill_styles(self, display: str):
        if self.engine == "openai" and not self._grouped:
            return
        name = display.rsplit(" (", 1)[0]
        styles = self._styles_by_speaker.get(name, [])
        self._value_by_label = dict(styles)
        labels = [lb for lb, _ in styles] or ["---"]
        self.style_menu.configure(values=labels)
        self.style_menu.set(labels[0])

    def _changed(self):
        self._pending = self.get()
        if self._on_change:
            self._on_change()

    def select(self, speaker: str, style: str = ""):
        """名前で選ぶ。一覧にない (または取得前の) ときは、一覧の最初の項目 (取得前は未取得の表示)。"""
        self._pending = (speaker, style)
        if not self._fetched:
            self._show_pending()
            return
        values = list(self.speaker_menu.cget("values") or [])
        if self.engine == "openai" and not self._grouped:
            self.speaker_menu.set(speaker if speaker in self._value_by_label else values[0])
            return
        if self.engine == "openai":
            # 声の ID から、その声を含むグループを選ぶ
            for display in values:
                styles = self._styles_by_speaker.get(display.rsplit(" (", 1)[0], [])
                lb = next((lb for lb, v in styles if v == speaker), None)
                if lb:
                    self.speaker_menu.set(display)
                    self._fill_styles(display)
                    self.style_menu.set(lb)
                    return
            self.speaker_menu.set(values[0])
            self._fill_styles(values[0])
            return
        display = next((d for d in values if d.rsplit(" (", 1)[0] == speaker), values[0])
        self.speaker_menu.set(display)
        self._fill_styles(display)
        lb = next((lb for lb in self._value_by_label if lb.rsplit(" (ID=", 1)[0] == style), None)
        if lb:
            self.style_menu.set(lb)

    def get(self) -> tuple[str, str]:
        """選択中の (話者名, スタイル名)。OpenAI互換は (声の ID, "")。取得前は選びたい名前。"""
        if not self._fetched:
            return self._pending
        if self.engine == "openai":
            if self._grouped:
                return str(self._value_by_label.get(self.style_menu.get(), "")), ""
            return self.speaker_menu.get(), ""
        speaker = self.speaker_menu.get()
        style = self.style_menu.get()
        return (speaker.rsplit(" (", 1)[0] if " (" in speaker else "",
                style.rsplit(" (ID=", 1)[0] if " (ID=" in style else "")

    def speakers(self) -> list[str]:
        """一覧の話者 (OpenAI互換は声の ID) を順に。役を追加するときに次の話者を選ぶため。"""
        if self.engine == "openai":
            if self._grouped:
                return [v for styles in self._styles_by_speaker.values() for _, v in styles]
            return list(self._value_by_label)
        return list(self._styles_by_speaker)
