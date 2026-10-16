"""その他タブ: アニメーション・ノートの確認・設定の保存・キャッシュ・更新の確認"""

import customtkinter as ctk

import i18n
import powerpoint
from cache import SynthCache
from i18n import t
from settings import Settings
from ui import widgets as w


def format_size(n: int) -> str:
    return f"{n / 1024 / 1024:.1f} MB" if n >= 1024 * 1024 else f"{n / 1024:.0f} KB"


class OtherTab(ctk.CTkScrollableFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self._build()

    def _build(self):
        # アニメーション
        w.heading(self, t("sec_animation"), pady=(8, 2))
        r = w.row(self)
        self.auto_next_enabled_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(r, text=t("auto_next"), variable=self.auto_next_enabled_var).pack(side="left")
        self.auto_next_var = ctk.DoubleVar(value=5.0)
        ctk.CTkEntry(r, textvariable=self.auto_next_var, width=50).pack(side="left", padx=(8, 0))
        w.hint(r, t("auto_next_interval")).pack(side="left", padx=(4, 0))
        w.hint(r, t("auto_next_off")).pack(side="left", padx=(8, 0))

        # ノートの確認
        w.heading(self, t("other_check"))
        r = w.row(self)
        ctk.CTkButton(r, text=t("check_notes"), width=140, command=self.app.check_notes).pack(side="left")
        w.hint(r, t("check_notes_desc")).pack(side="left", padx=(10, 0))

        # 設定の保存
        w.heading(self, t("save_config"))
        r = w.row(self)
        ctk.CTkButton(r, text=t("show_tags"), width=140, command=self.app.show_config_tags).pack(side="left")
        w.hint(r, t("show_tags_desc")).pack(side="left", padx=(10, 0))

        # キャッシュ (ユーザー設定。PPTX には保存しない)
        w.heading(self, t("other_cache"))
        r = w.row(self)
        self.use_cache_var = ctk.BooleanVar(value=bool(i18n.get_user_setting("use_cache", True)))
        ctk.CTkCheckBox(r, text=t("use_cache"), variable=self.use_cache_var,
                        command=lambda: i18n.set_user_setting("use_cache", self.use_cache_var.get())).pack(side="left")
        r = w.row(self)
        self.clear_cache_btn = ctk.CTkButton(r, text=t("clear_cache"), width=140, command=self._clear_cache)
        self.clear_cache_btn.pack(side="left")
        self.cache_size_label = w.hint(r, "")
        self.cache_size_label.pack(side="left", padx=(10, 0))
        w.hint(self, t("cache_desc"), justify="left").pack(fill="x", padx=12, pady=(0, 2))

        # PowerPoint との連携 (ユーザー設定)
        w.heading(self, t("other_powerpoint"))
        r = w.row(self)
        self.use_powerpoint_var = ctk.BooleanVar(value=bool(i18n.get_user_setting("use_powerpoint", True)))
        cb = ctk.CTkCheckBox(r, text=t("use_powerpoint"), variable=self.use_powerpoint_var,
                             command=self._on_use_powerpoint)
        cb.pack(side="left")
        if not powerpoint.is_available():
            cb.configure(state="disabled")
            w.hint(r, t("powerpoint_not_found")).pack(side="left", padx=(10, 0))
        r = w.row(self)
        w.label(r, t("video_height"))
        self.video_height_var = ctk.StringVar(value=f"{i18n.get_user_setting('video_height', 1080)}p")
        ctk.CTkComboBox(r, values=["480p", "720p", "1080p", "2160p"], variable=self.video_height_var, width=100,
                        state="readonly", command=lambda v: i18n.set_user_setting("video_height", int(v[:-1]))
                        ).pack(side="left", padx=(4, 0))
        w.hint(self, t("powerpoint_desc"), justify="left").pack(fill="x", padx=12, pady=(0, 2))

        # 更新の確認 (ユーザー設定)
        w.heading(self, t("other_updates"))
        r = w.row(self, pady=(2, 10))
        self.check_updates_var = ctk.BooleanVar(value=bool(i18n.get_user_setting("check_updates", True)))
        ctk.CTkCheckBox(r, text=t("check_updates"), variable=self.check_updates_var,
                        command=lambda: i18n.set_user_setting("check_updates", self.check_updates_var.get())
                        ).pack(side="left")
        self.refresh_cache_size()

    def refresh_cache_size(self):
        try:
            size = SynthCache().size()
        except OSError:
            size = 0
        self.cache_size_label.configure(text=t("cache_size", size=format_size(size)))

    def _clear_cache(self):
        try:
            SynthCache().clear()
            self.app._log(t("log_cache_cleared"))
        except OSError as e:
            self.app._log(t("log_error", e=e))
        self.refresh_cache_size()

    def _on_use_powerpoint(self):
        i18n.set_user_setting("use_powerpoint", self.use_powerpoint_var.get())
        self.app.apply_powerpoint_ui()

    @property
    def use_cache(self) -> bool:
        return self.use_cache_var.get()

    def load(self, s: Settings):
        self.auto_next_var.set(s.auto_next)
        self.auto_next_enabled_var.set(s.auto_next_enabled)

    def save(self, s: Settings):
        try:
            s.auto_next = float(self.auto_next_var.get())
        except Exception:
            s.auto_next = 5.0
        s.auto_next_enabled = self.auto_next_enabled_var.get()
