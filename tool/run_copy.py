"""Закреплённая копия сторожа почты для Планировщика (12.11, В5).

Боевой запуск идёт НЕ из рабочей копии разработчика (.tool_wt), а из отдельного git worktree .tool_run,
стоящего на одном коммите. Обновляется только этой командой, с явно названным коммитом или тегом:

    python tool\\run_copy.py <коммит|тег>                 — создать/передвинуть .tool_run на коммит
    python tool\\run_copy.py <коммит> --target <папка>    — другое место (для проверки)

Перед обновлением: копия чистая (нет незакоммиченных правок) и не идёт запуск (нет state\\mail_watch.lock).
При первом создании курсор и список «виденного» (state\\mail_watch.json) копируются из прежнего места запуска,
чтобы сторож не перечитал ящик с 09.11 заново. Ящик эта команда не открывает.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent


def git(*args, cwd=None, check=True) -> str:
    r = subprocess.run(["git", *args], cwd=cwd, capture_output=True, text=True, encoding="utf-8")
    if check and r.returncode != 0:
        raise SystemExit(f"git {' '.join(args)}: {r.stderr.strip()}")
    return r.stdout.strip()


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="закреплённая копия .tool_run для Планировщика")
    ap.add_argument("commit", help="коммит или тег, на который поставить копию")
    ap.add_argument("--repo", default=str(HERE.parent), help="любой worktree репозитория (по умолчанию этот)")
    ap.add_argument("--target", default=None, help="папка копии (по умолчанию <основной репозиторий>\\.tool_run)")
    ap.add_argument("--state-from", default=None, help="откуда взять state\\mail_watch.json при первом создании")
    a = ap.parse_args(argv)

    common = Path(git("rev-parse", "--path-format=absolute", "--git-common-dir", cwd=a.repo))
    main_dir = common.parent
    target = Path(a.target) if a.target else main_dir / ".tool_run"
    sha = git("rev-parse", "--verify", a.commit + "^{commit}", cwd=a.repo)
    state_from = Path(a.state_from) if a.state_from else main_dir / ".tool_wt" / "state" / "mail_watch.json"

    if target.exists():
        if not (target / ".git").exists():
            raise SystemExit(f"{target} есть, но это не git worktree — не трогаю")
        if (target / "state" / "mail_watch.lock").exists():
            raise SystemExit("идёт запуск сторожа (state\\mail_watch.lock) — обнови позже")
        dirty = git("status", "--porcelain", "--untracked-files=no", cwd=target)
        if dirty:
            raise SystemExit(f"в {target} незакоммиченные правки, не обновляю:\n{dirty}")
        before = git("rev-parse", "HEAD", cwd=target)
        git("checkout", "--detach", sha, cwd=target)
        print(f"{target}: {before[:7]} -> {sha[:7]}")
    else:
        git("worktree", "add", "--detach", str(target), sha, cwd=a.repo)
        print(f"{target}: создан на {sha[:7]}")
    st = target / "state" / "mail_watch.json"
    if not st.exists() and state_from.exists():
        st.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(state_from, st)
        print(f"состояние скопировано из {state_from}")
    excl = common / "info" / "exclude"                       # копия внутри папки основного репозитория: не в git status
    rel = target.resolve()
    try:
        name = rel.relative_to(main_dir.resolve()).as_posix() + "/"
        lines = excl.read_text(encoding="utf-8").splitlines() if excl.exists() else []
        if name not in lines:
            excl.parent.mkdir(parents=True, exist_ok=True)
            excl.write_text("\n".join(lines + [name]) + "\n", encoding="utf-8")
    except ValueError:
        pass                                                   # копия вне папки репозитория: исключать нечего
    print(f"боевой запуск: {target / 'tool' / 'mail_watch.py'} на {git('log', '-1', '--format=%h %s', cwd=target)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
