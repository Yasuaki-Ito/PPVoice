"""新しいバージョンの確認 (GitHub Releases の最新版と比べる)

ダウンロードやインストールはしない。失敗 (オフライン・レート制限・形式の違いなど) は何も知らせない。
"""

import re

import requests

RELEASES_API = "https://api.github.com/repos/Yasuaki-Ito/PPVoice/releases/latest"
TIMEOUT = (3, 5)

_VERSION = re.compile(r"v?(\d+)\.(\d+)\.(\d+)$")


def parse_version(text: str) -> tuple[int, int, int] | None:
    """"v0.7.2" / "0.7.2" → (0, 7, 2)。形式が違えば None。"""
    m = _VERSION.match(text.strip())
    return tuple(int(x) for x in m.groups()) if m else None


def newer_release(current: str, release: dict) -> tuple[str, str] | None:
    """リリースの情報 (GitHub API の応答) が今のバージョンより新しければ (タグ, リリースページの URL)。"""
    latest = parse_version(str(release.get("tag_name", "")))
    now = parse_version(current)
    if latest is None or now is None or latest <= now:
        return None
    return release["tag_name"], str(release.get("html_url", ""))


def check(current: str) -> tuple[str, str] | None:
    """最新のリリースを問い合わせ、新しいバージョンがあれば (タグ, URL)。なければ・失敗したら None。"""
    try:
        resp = requests.get(RELEASES_API, timeout=TIMEOUT, headers={"Accept": "application/vnd.github+json"})
        resp.raise_for_status()
        return newer_release(current, resp.json())
    except (requests.RequestException, ValueError, KeyError, AttributeError):
        return None
