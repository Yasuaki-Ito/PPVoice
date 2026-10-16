"""PPVoice のコマンドライン版 (AI エージェントなどからの実行を想定)

  python src/cli.py generate in.pptx [-o out.pptx] [--slides 1-3,5] [--force] [--set key=value ...] [--json]
  python src/cli.py check    in.pptx [--slides 1-3,5] [--set key=value ...] [--json]
  python src/cli.py speakers [--engine voicevox|openai] [--url URL] [--json]

- 設定は PPTX のノートの <config> を読み、--set で上書きする (キーは <config> と同じ。例: --set speed=1.2)
- 進捗とログは stderr、結果は stdout に出す。--json で機械が読める形にする
- 終了コード: 0 = 成功, 1 = 入力・設定のエラー, 2 = TTS サーバのエラー
- OpenAI 互換の API キーは環境変数 OPENAI_API_KEY から読む (コマンドの履歴に残らないように)
"""

import argparse
import contextlib
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import requests  # noqa: E402

import i18n  # noqa: E402
import pipeline  # noqa: E402
from lint import lint_notes  # noqa: E402
from notes import parse_notes  # noqa: E402
from pptx_reader import read_slides  # noqa: E402
from settings import DEFAULT_URLS, ENGINES, Settings, parse_config_tags  # noqa: E402
from version import __version__  # noqa: E402

EXIT_OK, EXIT_INPUT, EXIT_SERVER = 0, 1, 2


class CliError(Exception):
    def __init__(self, message: str, code: int = EXIT_INPUT):
        super().__init__(message)
        self.code = code


def parse_slides(spec: str | None) -> set[int] | None:
    """"1-3,5" → {1, 2, 3, 5}。None はすべてのスライド。"""
    if not spec:
        return None
    result: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            if "-" in part:
                a, b = part.split("-", 1)
                result.update(range(int(a), int(b) + 1))
            else:
                result.add(int(part))
        except ValueError:
            raise CliError(f"invalid --slides: {spec}")
    return result


def load_settings(args, input_path: str | None) -> Settings:
    """PPTX の <config> → --engine / --set / --url の順に反映した設定。"""
    settings = Settings()
    if input_path:
        settings.apply_config(pipeline.read_config(input_path),
                              warn=lambda k, v: print(i18n.t("log_config_invalid", key=k, value=v), file=sys.stderr))
        for name, values in pipeline.read_cast(input_path).items():
            settings.apply_cast({name: values}, warn=lambda k, v, name=name: print(
                i18n.t("log_cast_invalid", name=name, key=k, value=v), file=sys.stderr))
        settings.dictionary = pipeline.read_dict(input_path)
    if getattr(args, "engine", None):
        settings.set_engine(args.engine)
    overrides = {}
    for item in getattr(args, "set", None) or []:
        if "=" not in item:
            raise CliError(f"invalid --set (expected key=value): {item}")
        key, value = item.split("=", 1)
        overrides[key.strip()] = value.strip()
    if overrides:
        unknown = [k for k in overrides if k not in Settings.__dataclass_fields__ and k != "engine"]
        if unknown:
            raise CliError(f"unknown --set key: {', '.join(unknown)}")
        bad = []
        settings.apply_config(overrides, warn=lambda k, v: bad.append(f"{k}={v}"))
        if bad:
            raise CliError(f"invalid --set value: {', '.join(bad)}")
    if getattr(args, "url", None):
        settings.url = args.url
    settings.api_key = os.environ.get("OPENAI_API_KEY", "")
    return settings


def server_error(settings: Settings, e: Exception) -> CliError:
    """サーバへの接続・応答のエラーを、GUI と同じ文言の CliError にする。"""
    url = settings.url
    if isinstance(e, requests.exceptions.ConnectTimeout):
        msg = i18n.t("log_connect_timeout", url=url, sec=pipeline.FETCH_TIMEOUT[0])
    elif isinstance(e, requests.exceptions.ConnectionError):
        msg = i18n.t("log_connect_failed", url=url)
    elif isinstance(e, requests.exceptions.Timeout):
        msg = i18n.t("log_read_timeout", url=url, sec=pipeline.FETCH_TIMEOUT[1])
    else:
        msg = str(e)
    return CliError(msg.strip(), EXIT_SERVER)


def require_input(path: str):
    if not os.path.isfile(path):
        raise CliError(i18n.t("log_file_not_found", path=path).strip())


# ---------------------------------------------------------------------------
# サブコマンド
# ---------------------------------------------------------------------------

def cmd_generate(args) -> dict:
    require_input(args.input)
    output = args.output or os.path.splitext(args.input)[0] + "_speech.pptx"
    if os.path.exists(output) and not args.force:
        raise CliError(f"output already exists (use --force to overwrite): {output}")
    settings = load_settings(args, args.input)
    try:
        engine = pipeline.create_engine(settings)
        # ログは stderr へ (結果だけを stdout に出すため)
        with contextlib.redirect_stdout(sys.stderr):
            cast = pipeline.create_cast(settings)
            result = pipeline.generate(settings, args.input, output, selected_slides=parse_slides(args.slides),
                                       engine=engine, cast=cast)
    except requests.RequestException as e:
        raise server_error(settings, e)
    except ValueError as e:
        raise CliError(str(e), EXIT_SERVER)
    except Exception as e:
        # TTS サーバのエラー (TTSServerError など) は、メッセージをそのまま出す
        if type(e).__name__ == "TTSServerError":
            raise CliError(str(e).strip(), EXIT_SERVER)
        raise
    return {"output": result.output if result else None, "generated": result is not None,
            "engine": settings.engine, "voice": cast.label,
            "roles": result.roles if result else {}, "credits": result.credits if result else []}


def cmd_check(args) -> dict:
    require_input(args.input)
    settings = load_settings(args, args.input)
    selected = parse_slides(args.slides)
    slides = read_slides(args.input)
    next_checks = {c.slide: c for c in pipeline.check_next_tags(args.input)}
    default_role = pipeline.default_role_name(settings)
    # 話者の一覧 (サーバにつながる場合だけ)。スタイル名の確認にも使う
    try:
        speakers = pipeline.fetch_speakers(settings)
        server = True
    except requests.RequestException:
        speakers, server = None, False
    role_names = {default_role, *settings.cast}
    styles = None
    if server and settings.engine == "voicevox":
        by_speaker = {sp["name"]: {st["name"] for st in sp.get("styles", [])} for sp in speakers}
        styles = {default_role: by_speaker.get(settings.speaker, set())}
        styles.update({r.name: by_speaker.get(r.speaker, set()) for r in settings.cast.values()})
        styles = {name: st for name, st in styles.items() if st}
    roles: dict[str, int] = {}   # 役名 → 文の数 (ノートに出てくる順)
    dict_hits = {word: 0 for word in settings.dictionary}  # 辞書の語 → 当たった回数
    result_slides = []
    for si in slides:
        num = si.index + 1
        if selected is not None and num not in selected:
            continue
        parsed = parse_notes(si.notes_text, settings.split_sentences, settings.dictionary) if si.notes_text else None
        c = next_checks[num]
        sentences = parsed.sentences if parsed else []
        for word, n in (parsed.dict_hits if parsed else {}).items():
            dict_hits[word] += n
        for s in sentences:
            name = s.speaker or default_role
            roles[name] = roles.get(name, 0) + 1
        result_slides.append({
            "slide": num,
            "sentences": [{"subtitle": pipeline._unescape_display(s.display), "reading": s.reading,
                           "role": s.speaker or default_role, "style": s.style} for s in sentences],
            "animations": c.animations, "next": c.nexts, "next_status": c.status.strip(" ←"),
            "issues": [{"line": i.line, "kind": i.kind, "tag": i.text, "message": i.message}
                       for i in lint_notes(si.notes_text, role_names, default_role, styles, settings.engine)]
            if si.notes_text else [],
        })
    def available(speaker: str, style: str):
        if not server or not speaker:
            return None
        if settings.engine == "openai":
            return speaker in speakers
        names = {sp["name"]: [st["name"] for st in sp.get("styles", [])] for sp in speakers}
        return speaker in names and (not style or style in names[speaker])

    speaker_check = {"server": server, "speaker": settings.speaker, "style": settings.style,
                     "available": available(settings.speaker, settings.style)}
    # 役: 既定の役 + <cast> の役。ノートにあって定義されていない役は undefined_roles に
    cast = [{"name": default_role, "default": True, "speaker": settings.speaker, "style": settings.style,
             "available": speaker_check["available"], "sentences": roles.get(default_role, 0)}]
    cast += [{"name": r.name, "default": False, "speaker": r.speaker, "style": r.style,
              "available": available(r.speaker, r.style), "sentences": roles.get(r.name, 0)}
             for r in settings.cast.values() if r.name != default_role]
    defined = {c["name"] for c in cast}
    undefined = [name for name in roles if name not in defined]
    return {"engine": settings.engine, "url": settings.url, "config": parse_config_tags(
        [s.notes_text for s in slides if s.notes_text]), "speaker_check": speaker_check, "cast": cast,
        "undefined_roles": undefined, "dictionary": [{"word": w, "reading": r, "hits": dict_hits.get(w, 0)}
                                                      for w, r in settings.dictionary.items()],
        "slides": result_slides}


def cmd_speakers(args) -> dict:
    settings = load_settings(args, None)
    try:
        speakers = pipeline.fetch_speakers(settings)
    except requests.RequestException as e:
        raise server_error(settings, e)
    if settings.engine == "openai":
        return {"engine": "openai", "url": settings.url, "voices": speakers}
    return {"engine": "voicevox", "url": settings.url, "speakers": [
        {"name": sp["name"], "styles": [{"name": st["name"], "id": st["id"]} for st in sp.get("styles", [])]}
        for sp in speakers]}


# ---------------------------------------------------------------------------
# 出力
# ---------------------------------------------------------------------------

def print_text(command: str, result: dict):
    if command == "generate":
        print(result["output"] or "(no slides with notes)")
    elif command == "speakers":
        if "voices" in result:
            print("\n".join(result["voices"]))
        else:
            for sp in result["speakers"]:
                print(f"{sp['name']}: " + ", ".join(f"{st['name']} ({st['id']})" for st in sp["styles"]))
    elif command == "check":
        sc = result["speaker_check"]
        status = ("server not reachable" if not sc["server"] else
                  "ok" if sc["available"] else "not found" if sc["available"] is False else "not specified")
        print(f"engine: {result['engine']} ({result['url']}), speaker: {sc['speaker'] or '-'}"
              f"{' / ' + sc['style'] if sc['style'] else ''} [{status}]")
        if len(result["cast"]) > 1 or result["undefined_roles"]:
            for c in result["cast"]:
                ok = {True: "ok", False: "not found", None: "-"}[c["available"]]
                print(f"role {c['name'] or '-'}{' (default)' if c['default'] else ''}: {c['speaker'] or '-'}"
                      f"{' / ' + c['style'] if c['style'] else ''} [{ok}], {c['sentences']} sentences")
            for name in result["undefined_roles"]:
                print(f"role {name}: not defined (read by the default role)")
        for d in result["dictionary"]:
            print(f"dict {d['word']} = {d['reading']}: {d['hits']} hits")
        for sl in result["slides"]:
            note = f"  ({sl['next_status']})" if sl["next_status"] else ""
            print(f"--- slide {sl['slide']}: {len(sl['sentences'])} sentences, "
                  f"<next> {sl['next']} / animations {sl['animations']}{note}")
            multi = len(result["cast"]) > 1 or result["undefined_roles"]
            for issue in sl["issues"]:
                print(f"  ! line {issue['line']}: {issue['message']}")
            for s in sl["sentences"]:
                role = f"{s['role']}{'/' + s['style'] if s['style'] else ''}: " if multi else ""
                print(f"  {role}{s['subtitle']}"
                      + (f"  [reading: {s['reading']}]" if s["reading"] != s["subtitle"] else ""))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="ppvoice-cli", description="PPVoice command line interface")
    parser.add_argument("--version", action="version", version=f"PPVoice {__version__}")
    parser.add_argument("--lang", choices=list(i18n.LANGUAGES), help="language of log messages")
    sub = parser.add_subparsers(dest="command", required=True)

    def common(p, with_input=True):
        if with_input:
            p.add_argument("input", help="input PPTX")
        p.add_argument("--engine", choices=list(ENGINES), help="override the engine in <config>")
        p.add_argument("--url", help=f"TTS server URL (default: {DEFAULT_URLS})")
        p.add_argument("--set", action="append", metavar="KEY=VALUE",
                       help="override a setting, using the same keys as <config> (repeatable)")
        p.add_argument("--json", action="store_true", help="print the result as JSON")

    g = sub.add_parser("generate", help="generate a narrated PPTX")
    common(g)
    g.add_argument("-o", "--output", help="output PPTX (default: <input>_speech.pptx)")
    g.add_argument("--slides", help="slides to generate, e.g. 1-3,5")
    g.add_argument("--force", action="store_true", help="overwrite the output if it exists")

    c = sub.add_parser("check", help="parse the notes without synthesizing")
    common(c)
    c.add_argument("--slides", help="slides to check, e.g. 1-3,5")

    s = sub.add_parser("speakers", help="list speakers (voices for OpenAI-compatible servers)")
    common(s, with_input=False)
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    if args.lang:
        i18n.set_language(args.lang)
    else:
        i18n.load_language()
    try:
        result = {"generate": cmd_generate, "check": cmd_check, "speakers": cmd_speakers}[args.command](args)
    except CliError as e:
        if args.json:
            print(json.dumps({"error": str(e), "exit_code": e.code}, ensure_ascii=False))
        print(f"error: {e}", file=sys.stderr)
        return e.code
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=1))
    else:
        print_text(args.command, result)
    return EXIT_OK


if __name__ == "__main__":
    # Windows のコンソールでも日本語を出せるように
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    sys.exit(main())
