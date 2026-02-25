"""PowerPoint自動スピーチツール

PPTXのノート欄テキストから音声を合成し、音声付きPPTXを生成する。

使用例:
    # 音声付きPPTX生成
    python main.py input.pptx -o output.pptx

    # 字幕付き
    python main.py input.pptx -o output.pptx --subtitle

    # VOICEVOX話者IDを指定
    python main.py input.pptx -o output.pptx --speaker 3

    # 話者一覧を表示
    python main.py --list-speakers
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from version import __version__
from pptx_reader import read_slides
from pptx_writer import embed_audio
from tts.voicevox import VoicevoxEngine


def list_speakers(base_url: str) -> None:
    """VOICEVOX話者一覧を表示"""
    engine = VoicevoxEngine(base_url=base_url)
    try:
        speakers = engine.list_speakers()
    except Exception as e:
        print(f"VOICEVOXエンジンに接続できません: {e}", file=sys.stderr)
        print("VOICEVOXエンジンが起動しているか確認してください。", file=sys.stderr)
        sys.exit(1)

    for speaker in speakers:
        print(f"■ {speaker['name']}")
        for style in speaker.get("styles", []):
            print(f"    ID={style['id']:3d}  {style['name']}")


def generate_audio(slides, engine, with_timings=False):
    """各スライドのノートから音声を生成する。

    Returns:
        slide_audio: [(index, wav_bytes), ...]
        slide_timings: {index: [(text, start_ms, dur_ms), ...]}  (with_timings=True時)
    """
    slide_audio = []
    slide_timings = {}

    total_slides = len(slides)
    for info in slides:
        slide_num = info.index + 1
        if not info.notes_text:
            print(f"  [{slide_num}/{total_slides}] スライド {slide_num}: (ノートなし - スキップ)")
            slide_audio.append((info.index, b""))
            continue

        print(f"  [{slide_num}/{total_slides}] スライド {slide_num}:")

        def on_chunk(i, total, text, _sn=slide_num):
            print(f"    ({i + 1}/{total}) {text}")

        if with_timings:
            wav, timings = engine.synthesize_with_timings(info.notes_text, on_chunk=on_chunk)
            slide_timings[info.index] = timings
        else:
            wav = engine.synthesize(info.notes_text, on_chunk=on_chunk)
        slide_audio.append((info.index, wav))

    return slide_audio, slide_timings


def main():
    parser = argparse.ArgumentParser(
        description="PowerPointノートから自動スピーチを生成",
    )
    parser.add_argument("--version", action="version", version=f"PPVoice {__version__}")
    parser.add_argument("input", nargs="?", help="入力PPTXファイル")
    parser.add_argument("-o", "--output", help="出力ファイルパス")
    parser.add_argument("--speaker", type=int, default=1, help="VOICEVOX話者ID (default: 1)")
    parser.add_argument(
        "--voicevox-url",
        default="http://localhost:50021",
        help="VOICEVOX APIのURL (default: http://localhost:50021)",
    )
    parser.add_argument("--pause", type=float, default=0.5, help="文間の無音秒数 (default: 0.5)")
    parser.add_argument("--end-pause", type=float, default=2.0, help="スライド音声終了後の待機秒数 (default: 2.0)")
    parser.add_argument("--subtitle", action="store_true", help="字幕を表示する")
    parser.add_argument("--subtitle-style", choices=["box", "outline"], default="box",
                        help="字幕スタイル: box=半透明背景, outline=縁取り (default: box)")
    parser.add_argument("--subtitle-size", type=int, default=18, help="字幕フォントサイズ (default: 18)")
    parser.add_argument("--subtitle-font", default="", help="字幕のデフォルトフォント名 (default: テーマ依存)")
    parser.add_argument("--subtitle-bottom", type=float, default=0.05, help="字幕の下マージン (0.0〜1.0, default: 0.05)")
    parser.add_argument("--subtitle-color", default="FFFFFF", help="字幕テキスト色 hex RGB (default: FFFFFF)")
    parser.add_argument("--subtitle-outline", action=argparse.BooleanOptionalAction, default=True,
                        help="縁取りスタイル時に輪郭を付ける (default: on)")
    parser.add_argument("--subtitle-outline-color", default="000000", help="輪郭の色 hex RGB (default: 000000)")
    parser.add_argument("--subtitle-outline-width", type=float, default=0.75, help="輪郭の太さ pt (default: 0.75)")
    parser.add_argument("--subtitle-glow", action=argparse.BooleanOptionalAction, default=False,
                        help="縁取りスタイル時にぼかしを付ける (default: off)")
    parser.add_argument("--subtitle-glow-color", default="000000", help="ぼかしの色 hex RGB (default: 000000)")
    parser.add_argument("--subtitle-glow-size", type=float, default=11.0, help="ぼかしのサイズ pt (default: 11.0)")
    parser.add_argument("--subtitle-bg-color", default="000000", help="字幕背景色 hex RGB (default: 000000)")
    parser.add_argument("--subtitle-bg-alpha", type=int, default=60, help="字幕背景の不透明度 %% (0-100, default: 60)")
    parser.add_argument("--list-speakers", action="store_true", help="VOICEVOX話者一覧を表示")

    args = parser.parse_args()

    # 話者一覧モード
    if args.list_speakers:
        list_speakers(args.voicevox_url)
        return

    if not args.input:
        parser.error("入力PPTXファイルを指定してください")

    if not os.path.exists(args.input):
        print(f"ファイルが見つかりません: {args.input}", file=sys.stderr)
        sys.exit(1)

    # 出力パスのデフォルト設定
    base_name = os.path.splitext(args.input)[0]
    output_pptx = args.output if args.output else base_name + "_speech.pptx"

    # スライド読み込み
    print(f"PPTXを読み込んでいます: {args.input}")
    slides = read_slides(args.input)
    print(f"  {len(slides)} スライドを検出")

    notes_count = sum(1 for s in slides if s.notes_text)
    if notes_count == 0:
        print("ノートが含まれるスライドがありません。終了します。")
        return
    print(f"  {notes_count} スライドにノートあり")

    # 音声合成
    need_timings = args.subtitle
    print(f"\n音声を合成しています (speaker={args.speaker}, pause={args.pause}s)...")
    engine = VoicevoxEngine(speaker_id=args.speaker, base_url=args.voicevox_url, pause_sec=args.pause)
    try:
        slide_audio, slide_timings = generate_audio(slides, engine, with_timings=need_timings)
    except Exception as e:
        print(f"\n音声合成に失敗しました: {e}", file=sys.stderr)
        print("VOICEVOXエンジンが起動しているか確認してください。", file=sys.stderr)
        sys.exit(1)

    # PPTX出力
    print(f"\n音声付きPPTXを生成しています...")
    embed_audio(
        args.input,
        slide_audio,
        output_pptx,
        end_pause_ms=int(args.end_pause * 1000),
        slide_timings=slide_timings if need_timings else None,
        subtitle_font_size=args.subtitle_size,
        subtitle_font_name=args.subtitle_font,
        subtitle_bottom_pct=args.subtitle_bottom,
        subtitle_style=args.subtitle_style,
        subtitle_font_color=args.subtitle_color,
        subtitle_use_outline=args.subtitle_outline,
        subtitle_outline_color=args.subtitle_outline_color,
        subtitle_outline_width=args.subtitle_outline_width,
        subtitle_use_glow=args.subtitle_glow,
        subtitle_glow_color=args.subtitle_glow_color,
        subtitle_glow_size=args.subtitle_glow_size,
        subtitle_bg_color=args.subtitle_bg_color,
        subtitle_bg_alpha=args.subtitle_bg_alpha,
    )

    print("\n完了!")


if __name__ == "__main__":
    main()
