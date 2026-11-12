"""run_copy.py на временном репозитории: копия создаётся на коммите, состояние переносится, обновляется только явно."""
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run_copy  # noqa: E402


def g(cwd, *a):
    return subprocess.run(["git", *a], cwd=cwd, capture_output=True, text=True, check=True).stdout.strip()


@pytest.fixture
def repo(tmp_path):
    r = tmp_path / "repo"
    r.mkdir()
    g(r, "init", "-q")
    g(r, "config", "user.email", "t@example.com")
    g(r, "config", "user.name", "t")
    (r / "tool").mkdir()
    (r / "tool" / "mail_watch.py").write_text("v = 1\n", encoding="utf-8")
    (r / ".gitignore").write_text("state/\nlogs/\n", encoding="utf-8")
    g(r, "add", ".")
    g(r, "commit", "-q", "-m", "one")
    old_state = tmp_path / "old_state.json"
    old_state.write_text('{"folders": {"INBOX": {"last_uid": 26}}}', encoding="utf-8")
    return r, old_state


def test_creates_pinned_copy_with_state_and_moves_only_on_command(repo, tmp_path):
    r, old_state = repo
    first = g(r, "rev-parse", "HEAD")
    run = r / ".tool_run"
    assert run_copy.main([first, "--repo", str(r), "--state-from", str(old_state)]) == 0
    assert g(run, "rev-parse", "HEAD") == first
    assert '"last_uid": 26' in (run / "state" / "mail_watch.json").read_text(encoding="utf-8")
    assert ".tool_run/" in (r / ".git" / "info" / "exclude").read_text(encoding="utf-8")
    assert g(r, "status", "--porcelain") == ""                       # копия не видна в рабочей папке
    (r / "tool" / "mail_watch.py").write_text("v = 2\n", encoding="utf-8")
    g(r, "commit", "-q", "-am", "two")
    assert (run / "tool" / "mail_watch.py").read_text(encoding="utf-8") == "v = 1\n"   # новый коммит сам не приезжает
    second = g(r, "rev-parse", "HEAD")
    (run / "state" / "mail_watch.json").write_text('{"folders": {"INBOX": {"last_uid": 30}}}', encoding="utf-8")
    run_copy.main([second, "--repo", str(r), "--state-from", str(old_state)])
    assert (run / "tool" / "mail_watch.py").read_text(encoding="utf-8") == "v = 2\n"
    assert '"last_uid": 30' in (run / "state" / "mail_watch.json").read_text(encoding="utf-8")   # курсор копии не затёрт


def test_refuses_dirty_copy_and_running_lock(repo):
    r, old_state = repo
    head = g(r, "rev-parse", "HEAD")
    run_copy.main([head, "--repo", str(r), "--state-from", str(old_state)])
    run = r / ".tool_run"
    (run / "state" / "mail_watch.lock").write_text("123", encoding="utf-8")
    with pytest.raises(SystemExit, match="идёт запуск"):
        run_copy.main([head, "--repo", str(r)])
    (run / "state" / "mail_watch.lock").unlink()
    (run / "tool" / "mail_watch.py").write_text("hacked\n", encoding="utf-8")
    with pytest.raises(SystemExit, match="незакоммиченные"):
        run_copy.main([head, "--repo", str(r)])
