"""更新の確認 (update_check.py) のテスト。GitHub には接続しない。"""

import requests

import update_check
from update_check import newer_release, parse_version


def test_parse_version():
    assert parse_version("v0.7.2") == (0, 7, 2)
    assert parse_version("1.10.0") == (1, 10, 0)
    assert parse_version("v1.0.0-beta") is None


def test_newer_release():
    release = {"tag_name": "v1.0.0", "html_url": "https://example.com/r"}
    assert newer_release("0.7.3", release) == ("v1.0.0", "https://example.com/r")
    assert newer_release("1.0.0", release) is None
    assert newer_release("1.0.1", release) is None  # 開発中 (リリースより新しい)
    assert newer_release("0.7.3", {"tag_name": "nightly"}) is None


def test_check_is_silent_on_errors(monkeypatch):
    def fail(*a, **kw):
        raise requests.ConnectionError("offline")
    monkeypatch.setattr(update_check.requests, "get", fail)
    assert update_check.check("0.7.3") is None
