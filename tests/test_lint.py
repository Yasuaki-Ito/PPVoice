"""ノートの書き間違いの検出 (lint.py) のテスト"""

from lint import lint_notes


def kinds(text, **kw):
    return [(i.line, i.kind, i.text) for i in lint_notes(text, **kw)]


def test_valid_notes_have_no_issues():
    text = ("<config speaker=ずんだもん>\n<cast A speaker=四国めたん>\n<dict PPTX=パワーポイント>\n"
            "<speed=1.2><b>太字</b>と<color=#FF0000>赤</color>。<wait=1s><next>\n<size=+2>大</size><br>次<pitch=0.05>")
    assert kinds(text) == []


def test_unknown_tag_and_invalid_value():
    assert kinds("<sped=1.2>速く読む。\n<color=red>赤</color>\n<wait=abc>") == [
        (1, "unknown_tag", "<sped=1.2>"), (2, "invalid_value", "<color=red>"), (3, "invalid_value", "<wait=abc>")]


def test_tags_inside_braces_are_text():
    assert kinds("{<sped>|すぺっど} と {<foo>} と {$a<b>c$|えー}") == []


def test_unclosed_format_within_sentence():
    assert kinds("<b>太字\nここまで</b>") == [(1, "unclosed_format", "<b>"), (2, "unclosed_format", "<b>")]
    # <wait> と <speaker> で文が分かれる
    assert kinds("<i>a<wait=1s>b</i>") == [(1, "unclosed_format", "<i>"), (1, "unclosed_format", "<i>")]
    assert kinds("<u>a</u><speaker=A><b>b</b>") == []


def test_undefined_role():
    text = "<speaker=A>一。\n<speaker=B>二。\n<speaker=ずんだもん>三。"
    assert kinds(text, roles={"A", "ずんだもん"}) == [(2, "undefined_role", "<speaker=B>")]
    assert kinds(text) == []  # 役を渡さなければ確認しない


def test_style_checked_against_the_current_role():
    styles = {"ずんだもん": {"ノーマル", "ささやき"}, "めたん": {"ノーマル", "あまあま"}}
    text = "<style=ささやき>一。\n<speaker=めたん><style=ささやき>二。\n<style=あまあま>三。"
    issues = lint_notes(text, roles={"ずんだもん", "めたん"}, default_role="ずんだもん", styles=styles)
    assert [(i.line, i.kind, i.role, i.style) for i in issues] == [(2, "unknown_style", "めたん", "ささやき")]


def test_style_ignored_on_openai():
    assert kinds("<style=x>Hi.", engine="openai") == [(1, "style_ignored", "<style=x>")]
