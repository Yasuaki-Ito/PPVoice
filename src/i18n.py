"""GUI・ログの多言語化 (日本語 / 英語)

文言は STRINGS に「キー: (日本語, 英語)」の対で定義し、t("キー") で現在の言語の文言を引く。
言語はユーザーごとの設定として %APPDATA%\\PPVoice\\settings.json に保存する
(PPTX ごとの <config> タグには入れない)。初回は OS の表示言語から決める。
"""

import ctypes
import json
import os

LANGUAGES = {"ja": "日本語", "en": "English"}

_lang = "ja"

STRINGS: dict[str, tuple[str, str]] = {
    # --- 共通 ---
    "browse": ("参照", "Browse"),

    # --- ファイル設定 ---
    "sec_file": ("ファイル設定", "Files"),
    "input_file": ("入力ファイル (.pptx)", "Input file (.pptx)"),
    "output_file": ("出力ファイル (.pptx)", "Output file (.pptx)"),
    "slides": ("スライド", "Slides"),
    "slides_all": ("全部", "All"),
    "slides_some": ("一部", "Selected"),
    "slides_select": ("選択...", "Select..."),

    # --- 音声設定 ---
    "sec_voice": ("音声設定", "Voice"),
    "engine": ("エンジン", "Engine"),
    "engine_openai": ("OpenAI互換 (Kokoro等)", "OpenAI-compatible (Kokoro etc.)"),
    "voicevox_url": ("VOICEVOX URL", "VOICEVOX URL"),
    "url": ("URL", "URL"),
    "model": ("モデル", "Model"),
    "api_key": ("APIキー", "API key"),
    "api_key_placeholder": ("(不要なら空欄)", "(leave blank if not needed)"),
    "speaker": ("話者", "Speaker"),
    "voice": ("声", "Voice"),
    # Kokoro の声のグループ (言語・性別)
    "voice_group": ("言語・性別", "Language"),
    "voice_group_label": ("{lang}・{gender}", "{lang}, {gender}"),
    "voice_group_other": ("その他", "Other"),
    "gender_f": ("女性", "female"),
    "gender_m": ("男性", "male"),
    "lang_unknown": ("{code}（不明）", "{code} (unknown)"),
    "kokoro_lang_a": ("アメリカ英語", "American English"),
    "kokoro_lang_b": ("イギリス英語", "British English"),
    "kokoro_lang_e": ("スペイン語", "Spanish"),
    "kokoro_lang_f": ("フランス語", "French"),
    "kokoro_lang_h": ("ヒンディー語", "Hindi"),
    "kokoro_lang_i": ("イタリア語", "Italian"),
    "kokoro_lang_j": ("日本語", "Japanese"),
    "kokoro_lang_p": ("ポルトガル語 (ブラジル)", "Brazilian Portuguese"),
    "kokoro_lang_z": ("中国語", "Mandarin Chinese"),
    "fetch_speakers": ("話者取得", "Get speakers"),
    "fetch_voices": ("声を取得", "Get voices"),
    "fetching": ("取得中...", "Loading..."),
    "press_fetch": ("(取得ボタンを押してください)", "(Click the Get button)"),
    "style": ("スタイル", "Style"),
    "speed": ("速度", "Speed"),
    "pitch": ("ピッチ", "Pitch"),
    "intonation": ("抑揚", "Intonation"),
    "volume": ("音量", "Volume"),
    "test": ("テスト", "Test"),
    "test_text": ("音声のテストです。", "This is a voice test."),
    "play": ("▶ 再生", "▶ Play"),
    "stop_play": ("■ 停止", "■ Stop"),
    "synthesizing": ("合成中...", "Synthesizing..."),
    "voicevox_terms": ("※ キャラクターごとに利用規約があります → VOICEVOX公式サイト",
                       "* Each character has its own terms of use → VOICEVOX website"),
    "pause": ("文の区切り (秒)", "Sentence gap (s)"),
    "end_pause": ("末尾の余白 (秒)", "End padding (s)"),
    "split_sentences": ("文末 (. ! ?) でも区切る", "Also split at . ! ?"),

    # --- 字幕設定 ---
    "sec_subtitle": ("字幕設定", "Subtitles"),
    "show_subtitles": ("字幕を表示する", "Show subtitles"),
    "punct_replace": ("置換", "Replace"),
    "touten": ("読点", "Comma"),
    "kuten": ("句点", "Period"),
    "decoration": ("文字装飾", "Text style"),
    "bold": ("太字", "Bold"),
    "italic": ("斜体", "Italic"),
    "underline": ("下線", "Underline"),
    "math_bold": ("数式を太字", "Bold math"),
    "sub_style": ("スタイル", "Style"),
    "style_outline": ("縁取り", "Outline"),
    "style_box": ("背景付き", "Background"),
    "font_size": ("フォントサイズ", "Font size"),
    "font": ("フォント", "Font"),
    "font_default": ("<テーマのデフォルト>", "<Theme default>"),
    "select": ("選択", "Select"),
    "bottom_margin": ("下マージン", "Bottom margin"),
    "font_color": ("文字色", "Text color"),
    "outline": ("輪郭", "Outline"),
    "outline_width": ("太さ", "Width"),
    "glow": ("ぼかし", "Glow"),
    "glow_size": ("サイズ", "Size"),
    "bg_color": ("背景色", "Background"),
    "bg_alpha": ("不透明度", "Opacity"),
    # 句読点の置換先 (内部の値は日本語の表記のまま。表示だけ翻訳する)
    "punct_unchanged": ("そのまま", "Unchanged"),
    "punct_comma_half": (",(半角)", ", (half-width)"),
    "punct_comma_full": ("，(全角)", "， (full-width)"),
    "punct_period_half": (".(半角)", ". (half-width)"),
    "punct_period_full": ("．(全角)", "． (full-width)"),
    "punct_space_half": ("(半角空白)", "(half-width space)"),
    "punct_space_full": ("(全角空白)", "(full-width space)"),

    # --- アニメーション ---
    "sec_animation": ("アニメーション", "Animation"),
    "auto_next": ("未指定アニメを自動再生", "Auto-play extra animations"),
    "auto_next_interval": ("秒間隔", "s interval"),
    "auto_next_off": ("(OFFでクリック待ち)", "(off: wait for click)"),
    "check_next": ("<next>確認", "Check <next>"),

    # --- 実行 ---
    "save_config": ("設定保存", "Save settings"),
    "run": ("生成開始", "Generate"),
    "stop": ("停止", "Stop"),
    "stopping": ("停止中...", "Stopping..."),

    # --- ダイアログ ---
    "slide_select_title": ("スライド選択", "Select slides"),
    "slide_select_heading": ("作成するスライドを選択 ({total}枚)", "Select slides to generate ({total} slides)"),
    "select_all": ("全選択", "Select all"),
    "deselect_all": ("全解除", "Clear all"),
    "slide_n": ("スライド {num}", "Slide {num}"),
    "no_selection_title": ("選択なし", "No selection"),
    "no_selection": ("少なくとも1枚選択してください。", "Select at least one slide."),
    "save_config_desc": ("以下のタグを PPTX の任意のスライドのノート欄に貼り付けると、\n"
                         "ファイルを開いた際に設定が自動的に読み込まれます。",
                         "Paste this tag into the notes of any slide in the PPTX.\n"
                         "The settings are loaded automatically when you open the file."),
    "copy": ("コピー", "Copy"),
    "copied": ("コピーしました", "Copied"),
    "font_picker_title": ("フォント選択", "Select font"),
    "search": ("検索...", "Search..."),
    "color_picker_title": ("色を選択", "Select color"),
    "overwrite_title": ("上書き確認", "Confirm overwrite"),
    "overwrite_msg": ("以下のファイルが既に存在します。上書きしますか?\n\n{names}",
                      "The following file already exists. Overwrite it?\n\n{names}"),

    # --- ログ: ファイル・スライド ---
    "log_config_loaded": ("設定タグを読み込みました:\n{details}\n", "Loaded settings from the <config> tag:\n{details}\n"),
    "log_need_input_for_select": ("スライド選択にはまず入力ファイルを指定してください。\n",
                                  "Specify an input file before selecting slides.\n"),
    "log_read_failed": ("ファイル読み込み失敗: {e}\n", "Failed to read the file: {e}\n"),
    "log_no_slides": ("スライドが見つかりませんでした。\n", "No slides were found.\n"),
    "log_no_input": ("入力ファイルが指定されていません。\n", "No input file is specified.\n"),
    "log_pptx_read_failed": ("PPTXの読み込みに失敗: {e}\n", "Failed to read the PPTX: {e}\n"),
    "log_specify_input": ("入力ファイルを指定してください。\n", "Specify an input file.\n"),
    "log_file_not_found": ("ファイルが見つかりません: {path}\n", "File not found: {path}\n"),

    # --- ログ: 話者取得 ---
    "engine_name_voicevox": ("VOICEVOXエンジン", "VOICEVOX engine"),
    "engine_name_server": ("TTSサーバ", "TTS server"),
    "log_connecting": ("{name} ({url}) に接続中...\n", "Connecting to {name} ({url})...\n"),
    "log_connect_timeout": ("話者取得失敗: {url} に接続できませんでした (タイムアウト {sec}秒)。",
                            "Could not get speakers: could not connect to {url} (timed out after {sec} s)."),
    "log_connect_failed": ("話者取得失敗: {url} に接続できませんでした。",
                           "Could not get speakers: could not connect to {url}."),
    "log_read_timeout": ("話者取得失敗: {url} から応答がありませんでした (タイムアウト {sec}秒)。",
                         "Could not get speakers: no response from {url} (timed out after {sec} s)."),
    "log_fetch_failed": ("話者取得失敗: {e}", "Could not get speakers: {e}"),
    "log_check_server": ("\n{name}が起動しているか、URLが正しいか確認してください。\n",
                         "\nCheck that the {name} is running and the URL is correct.\n"),
    "log_voices_fetched": ("{n} 個の声を取得しました。\n", "Voices found: {n}\n"),
    "log_no_voices": ("声が見つかりませんでした。\n", "No voices were found.\n"),
    "log_speakers_fetched": ("{n} 話者 ({styles} スタイル) を取得しました。\n",
                             "Speakers found: {n} ({styles} styles)\n"),
    "log_no_speakers": ("話者が見つかりませんでした。\n", "No speakers were found.\n"),
    "log_fetch_first": ("先に話者を取得してください。\n", "Get the speakers first.\n"),

    # --- ログ: <next> チェック ---
    "log_next_check_header": ("--- <next> / アニメーション チェック ---\n", "--- <next> / animation check ---\n"),
    "next_too_many": (" ← <next> が多い (余分は無視)", " ← more <next> than animations (extra ones are ignored)"),
    "next_no_anim": (" ← アニメなし (<next> は無視)", " ← no animations (<next> is ignored)"),
    "next_surplus": (" ← 余り {n} グループ", " ← extra animation groups: {n}"),
    "next_none": (" ← <next> なし (すべて未指定アニメとして扱う)", " ← no <next> (all animations are treated as extra)"),
    "log_next_check_row": ("  スライド {num}: アニメ={anims}, <next>={nexts}{status}\n",
                           "  Slide {num}: animations={anims}, <next>={nexts}{status}\n"),

    # --- ログ: テスト再生 ---
    "log_subtitle_preview": ("--- 字幕プレビュー ---\n", "--- Subtitle preview ---\n"),
    "log_test_error": ("テスト再生エラー: {msg}\n", "Test playback error: {msg}\n"),

    # --- ログ: 生成 ---
    "log_cancelled": ("\n処理を中断しました。", "\nCancelled."),
    "log_error": ("\nエラー: {e}", "\nError: {e}"),
    "log_reading_pptx": ("PPTXを読み込んでいます: {path}", "Reading PPTX: {path}"),
    "log_slides_found": ("  {n} スライドを検出", "  Slides found: {n}"),
    "log_slides_selected": ("  {n} スライドを選択中", "  Slides selected: {n}"),
    "log_no_notes": ("ノートが含まれるスライドがありません。終了します。", "No slides have notes. Nothing to do."),
    "log_notes_count": ("  {n} スライドにノートあり", "  Slides with notes: {n}"),
    "log_synthesizing": ("\n音声を合成しています ({desc}, pause={pause}s)...", "\nSynthesizing speech ({desc}, pause={pause}s)..."),
    "log_slide_skip": ("  [{i}/{total}] スライド {num}: (ノートなし - スキップ)", "  [{i}/{total}] Slide {num}: (no notes - skipped)"),
    "log_slide": ("  [{i}/{total}] スライド {num}:", "  [{i}/{total}] Slide {num}:"),
    "log_embedding": ("\n音声付きPPTXを生成しています...", "\nCreating the narrated PPTX..."),
    "log_done": ("\n完了! → {name}", "\nDone! → {name}"),
    "log_video_howto": ("\n--- 動画 (MP4) にするには ---\n"
                        "1. 生成されたPPTXをPowerPointで開く\n"
                        "2. ファイル → エクスポート → ビデオの作成\n"
                        "3. 品質を選択して「ビデオの作成」をクリック",
                        "\n--- To create a video (MP4) ---\n"
                        "1. Open the generated PPTX in PowerPoint\n"
                        "2. File → Export → Create a Video\n"
                        "3. Choose the quality and click \"Create Video\""),

    # --- ログ: 埋め込み・エンジン ---
    "log_saved": ("音声付きPPTX を保存しました: {path}", "Saved the narrated PPTX: {path}"),
    "log_math_failed": ("数式の変換に失敗しました (LaTeX のまま表示します): {latex} ({e})",
                        "Could not convert the math (shown as LaTeX text): {latex} ({e})"),
    "log_accent_failed": ("[PPVoice] アクセント再計算に失敗しました (mora_pitch/mora_data 未対応)",
                          "[PPVoice] Could not recalculate the accent (mora_pitch/mora_data not supported)"),
    "log_unsupported": ("[PPVoice] このエンジンでは {items} は使えないため無視します",
                        "[PPVoice] {items} is not supported by this engine and is ignored"),
    "accent_tag": ("アクセント指定 {…|…|N}", "accent {…|…|N}"),
    "log_server_error": ("TTSサーバがエラーを返しました (HTTP {status}): {message}\n  読み上げようとした文: {text}",
                         "The TTS server returned an error (HTTP {status}): {message}\n  Sentence: {text}"),
    "hint_no_speakable": ("\n  → 読み上げられる文字がないと判断されました。声の言語とノートの言語が合っているか確認してください。"
                          "日本語の声では、サーバ側で日本語の処理 (MeCab の辞書など) が必要な場合があります。",
                          "\n  → The server found nothing to read. Check that the voice language matches the notes. "
                          "Japanese voices may need Japanese text processing (such as a MeCab dictionary) on the server."),
    "hint_server_log": ("\n  詳しい原因は TTS サーバのログを確認してください。\n",
                        "\n  See the TTS server log for details.\n"),
}


def t(key: str, **kwargs) -> str:
    """現在の言語の文言を返す。kwargs があれば str.format で埋め込む。"""
    ja, en = STRINGS[key]
    s = en if _lang == "en" else ja
    return s.format(**kwargs) if kwargs else s


def get_language() -> str:
    return _lang


def set_language(lang: str):
    global _lang
    if lang in LANGUAGES:
        _lang = lang


def detect_os_language() -> str:
    """OS の表示言語が日本語なら "ja"、それ以外は "en" を返す。"""
    try:
        lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
        return "ja" if (lang_id & 0x3FF) == 0x11 else "en"  # 0x11 = LANG_JAPANESE
    except Exception:
        return "ja"


# ---------------------------------------------------------------------------
# ユーザー設定 (%APPDATA%\PPVoice\settings.json)
# ---------------------------------------------------------------------------

def _settings_path() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "PPVoice", "settings.json")


def load_language():
    """保存された言語を読み込む。なければ OS の表示言語を使う。"""
    try:
        with open(_settings_path(), encoding="utf-8") as f:
            lang = json.load(f).get("language")
    except (OSError, ValueError):
        lang = None
    set_language(lang if lang in LANGUAGES else detect_os_language())


def save_language(lang: str):
    path = _settings_path()
    try:
        settings = {}
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                settings = json.load(f)
        settings["language"] = lang
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except (OSError, ValueError):
        pass
