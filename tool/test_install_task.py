"""install_task.py: шаблон без личных данных, подстановка выдуманных значений, вызов schtasks подменён."""
import os
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import install_task as it  # noqa: E402

NS = {"t": "http://schemas.microsoft.com/windows/2004/02/mit/task"}
FAKE = dict(userid="PC\\user", pythonw="D:\\Py\\pythonw.exe", tool_run="D:\\repo\\.tool_run")


def test_template_has_only_marks_no_personal_data():
    t = it.TEMPLATE.read_text(encoding="utf-8")
    for m in it.MARKS:
        assert m in t
    # имя компьютера и пользователя берём из окружения, чтобы их не было в самом тесте
    bad = [os.environ.get("USERDOMAIN", ""), os.environ.get("USERNAME", ""), "Users" + chr(92), "AppData", ".tool_wt"]
    assert not [b for b in bad if b and b.lower() in t.lower()]
    assert not (it.TEMPLATE.parent / "EAA_mail_watch.xml").exists()          # готового XML в репозитории нет


def test_render_substitutes_fake_values_and_keeps_schedule(tmp_path):
    out = tmp_path / "task.xml"
    calls = []
    rc = it.main(["--userid", FAKE["userid"], "--pythonw", FAKE["pythonw"], "--tool-run", FAKE["tool_run"], "--out", str(out), "--dry"],
                 runner=lambda *a, **k: calls.append(a))
    assert rc == 0 and calls == []                                            # --dry: schtasks не вызывается
    raw = out.read_bytes()
    assert raw[:2] == b"\xff\xfe"
    text = raw[2:].decode("utf-16-le")
    assert "__" not in text.replace("__init__", "")
    r = ET.parse(out).getroot()
    assert r.find(".//t:LogonTrigger/t:UserId", NS).text == "PC\\user" and r.find(".//t:Principal/t:UserId", NS).text == "PC\\user"
    assert r.find(".//t:Exec/t:Command", NS).text == "D:\\Py\\pythonw.exe"
    assert r.find(".//t:Exec/t:Arguments", NS).text == '"D:\\repo\\.tool_run\\tool\\mail_watch.py"'
    assert r.find(".//t:Exec/t:WorkingDirectory", NS).text == "D:\\repo\\.tool_run\\tool"
    cal = r.find(".//t:CalendarTrigger", NS)
    assert cal.find("t:Repetition/t:Interval", NS).text == "PT1H" and cal.find("t:Repetition/t:Duration", NS).text == "PT15H30M"
    assert cal.find("t:StartBoundary", NS).text.endswith("T08:00:00")
    assert r.find(".//t:StartWhenAvailable", NS).text == "true" and r.find(".//t:MultipleInstancesPolicy", NS).text == "IgnoreNew"


def test_install_calls_schtasks_with_rendered_xml(tmp_path):
    run = tmp_path / ".tool_run" / "tool"
    run.mkdir(parents=True)
    (run / "mail_watch.py").write_text("", encoding="utf-8")
    pyw = tmp_path / "pythonw.exe"
    pyw.write_bytes(b"")
    out = tmp_path / "out.xml"
    calls = []

    class R:
        returncode, stdout, stderr = 0, "SUCCESS", ""

    rc = it.main(["--userid", "PC\\user", "--pythonw", str(pyw), "--tool-run", str(tmp_path / ".tool_run"), "--out", str(out)],
                 runner=lambda args, **k: calls.append(args) or R())
    assert rc == 0 and calls == [["schtasks", "/create", "/xml", str(out), "/tn", "\\EAA\\mail_watch", "/f"]]


def test_refuses_empty_userid_and_missing_copy(tmp_path):
    with pytest.raises(SystemExit, match="userid"):
        it.render(it.TEMPLATE.read_text(encoding="utf-8"), "", "x", "y")
    with pytest.raises(SystemExit, match="закреплённой копии"):
        it.main(["--userid", "PC\\user", "--pythonw", sys.executable, "--tool-run", str(tmp_path / "nope"), "--out", str(tmp_path / "o.xml")],
                runner=lambda *a, **k: None)
