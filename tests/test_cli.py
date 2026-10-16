"""cli.py の引数処理のテスト (TTS サーバには接続しない)"""

import argparse
import json

import pytest

import cli


def test_parse_slides():
    assert cli.parse_slides(None) is None
    assert cli.parse_slides("1-3,5") == {1, 2, 3, 5}
    assert cli.parse_slides(" 2 , 4-4 ") == {2, 4}
    with pytest.raises(cli.CliError):
        cli.parse_slides("a-b")


def args(**kw):
    defaults = dict(engine=None, set=None, url=None)
    defaults.update(kw)
    return argparse.Namespace(**defaults)


def test_load_settings_set_overrides(monkeypatch):
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    s = cli.load_settings(args(set=["speed=1.2", "font=Arial Rounded MT Bold", "subtitle=off"]), None)
    assert (s.speed, s.font, s.subtitle) == (1.2, "Arial Rounded MT Bold", False)


def test_load_settings_engine_and_url(monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    s = cli.load_settings(args(engine="openai", url="http://example:8880/v1"), None)
    assert (s.engine, s.url, s.api_key, s.split_sentences) == ("openai", "http://example:8880/v1", "sk-test", True)


@pytest.mark.parametrize("item", ["speed", "unknown_key=1", "speed=fast"])
def test_load_settings_rejects_bad_set(item):
    with pytest.raises(cli.CliError):
        cli.load_settings(args(set=[item]), None)


def test_missing_input_is_input_error(capsys, tmp_path):
    code = cli.main(["check", str(tmp_path / "none.pptx"), "--json"])
    assert code == cli.EXIT_INPUT
    assert json.loads(capsys.readouterr().out)["exit_code"] == cli.EXIT_INPUT


def test_existing_output_needs_force(capsys, tmp_path):
    src = tmp_path / "in.pptx"
    src.write_bytes(b"dummy")
    (tmp_path / "in_speech.pptx").write_bytes(b"exists")
    assert cli.main(["generate", str(src)]) == cli.EXIT_INPUT
    assert "--force" in capsys.readouterr().err
