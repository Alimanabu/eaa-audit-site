"""Установка задачи Планировщика сторожа почты из шаблона (12.11, reviewer, круг 3).

В репозитории лежит только шаблон deploy\\tasks\\EAA_mail_watch.template.xml с метками __USERID__, __PYTHONW__,
__TOOL_RUN__ — без имени компьютера, пользователя и путей его профиля. Эта команда подставляет значения этого
компьютера, пишет готовый XML (UTF-16) в deploy\\out\\ (папка в .gitignore) и регистрирует задачу:

    python tool\\install_task.py                 — подставить, записать и schtasks /create /xml … /tn \\EAA\\mail_watch /f
    python tool\\install_task.py --dry           — только записать готовый XML и показать путь, без schtasks
    --tool-run <папка>  (по умолчанию <основной репозиторий>\\.tool_run)   --pythonw <путь>   --userid <ДОМЕН\\имя>
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE.parent / "deploy" / "tasks" / "EAA_mail_watch.template.xml"
MARKS = ("__USERID__", "__PYTHONW__", "__TOOL_RUN__")


def render(template: str, userid: str, pythonw: str, tool_run: str) -> str:
    for name, val in (("userid", userid), ("pythonw", pythonw), ("tool_run", tool_run)):
        if not val or "__" in val:
            raise SystemExit(f"пустое или неверное значение {name}: {val!r}")
    out = (template.replace("__USERID__", escape(userid)).replace("__PYTHONW__", escape(pythonw))
           .replace("__TOOL_RUN__", escape(str(tool_run).rstrip("\\/"))))
    left = [m for m in MARKS if m in out]
    if left:
        raise SystemExit(f"в XML остались метки: {left}")
    return out


def default_values() -> dict:
    common = subprocess.run(["git", "rev-parse", "--path-format=absolute", "--git-common-dir"], cwd=HERE,
                            capture_output=True, text=True, encoding="utf-8").stdout.strip()
    main_dir = Path(common).parent if common else HERE.parent
    dom, user = os.environ.get("USERDOMAIN", ""), os.environ.get("USERNAME", "")
    return dict(userid=f"{dom}\\{user}" if dom and user else "",
                pythonw=str(Path(sys.executable).with_name("pythonw.exe")),
                tool_run=str(main_dir / ".tool_run"))


def main(argv=None, runner=subprocess.run) -> int:
    ap = argparse.ArgumentParser(description="установка задачи \\EAA\\mail_watch из шаблона")
    ap.add_argument("--userid"), ap.add_argument("--pythonw"), ap.add_argument("--tool-run")
    ap.add_argument("--out", default=str(HERE.parent / "deploy" / "out" / "EAA_mail_watch.xml"))
    ap.add_argument("--tn", default="\\EAA\\mail_watch")
    ap.add_argument("--dry", action="store_true", help="только записать XML, без schtasks")
    a = ap.parse_args(argv)
    d = default_values()
    userid, pythonw, tool_run = a.userid or d["userid"], a.pythonw or d["pythonw"], a.tool_run or d["tool_run"]
    if not a.dry:
        if not Path(pythonw).exists():
            raise SystemExit(f"нет {pythonw}")
        if not (Path(tool_run) / "tool" / "mail_watch.py").exists():
            raise SystemExit(f"нет закреплённой копии {tool_run}: сначала python tool\\run_copy.py <коммит>")
    xml = render(TEMPLATE.read_text(encoding="utf-8"), userid, pythonw, tool_run)
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(b"\xff\xfe" + xml.replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-16-le"))
    print(f"готовый XML: {out}")
    if a.dry:
        return 0
    r = runner(["schtasks", "/create", "/xml", str(out), "/tn", a.tn, "/f"], capture_output=True, text=True)
    print((r.stdout or "") + (r.stderr or ""))
    return r.returncode


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    sys.exit(main())
