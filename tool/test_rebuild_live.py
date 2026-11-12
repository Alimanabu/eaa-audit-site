"""rebuild_live.py на подменённом ящике: --dry ничего не пишет, --apply снимает ложные тревоги и пишет «не писать»."""
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import mail_watch as mw  # noqa: E402
import rebuild_live as rb  # noqa: E402
from test_mail_watch import (FakeIMAP, PW, dmarc_from, dsn, entry, env, promo, put_live, rows,  # noqa: E402,F401
                             state_12_11)


def setup(env):
    """Как 12.11 17:35: курсор INBOX 3, мягкая метка у мёртвого адреса, 4 тревоги в eaa_mail.json."""
    cfg = env
    raws = [dsn("contact@dead-shop.example", status="5.0.0", diag="smtp; 550 No Such User Here"),
            dmarc_from("MAILER-DAEMON@esa2.mx-relay.example", mid="<rid-0001@mx-relay>"),
            promo(),
            dsn("info@live-shop.example", status="5.0.0", diag="smtp; 550 No Such User Here")]     # UID 4 — после курсора
    mw._write_json(cfg.state_json, {"live": [], "folders": {"INBOX": {"uidvalidity": 111, "last_uid": 3}}})
    import openpyxl
    wb = openpyxl.load_workbook(cfg.sent_xlsx)
    wb.active.cell(row=2, column=6).value = "отправлено; недоставка временно/отказ"
    wb.save(cfg.sent_xlsx)
    put_live(cfg, state_12_11())
    return raws


def test_dry_writes_nothing_and_shows_3_alerts_to_0(env):
    raws = setup(env)
    before = {p: p.read_bytes() for p in (env.sent_xlsx, env.live_json, env.state_json)}
    box = FakeIMAP(raws)
    lines = []
    res = rb.rebuild(env, apply=False, imap_factory=lambda c, p: box, password=PW, out=lines.append)
    assert {p: p.read_bytes() for p in before} == before and not env.stop_xlsx.exists()      # ничего не записано
    assert box.readonly and all("BODY.PEEK" in a[1] for c, a in box.cmds if c == "FETCH")
    assert res["letters"] == 3                                                                  # UID 4 после курсора не взят
    assert res["alerts_before"] == 4 and res["alerts_after"] == 0 and res["file_alert_after"] is False
    assert [r[1] for r in res["stop_added"]] == ["contact@dead-shop.example"]
    assert res["marks_changed"]["contact@dead-shop.example"][1] == "отправлено; нерабочий"
    assert any("--dry" in ln for ln in lines) and PW not in "\n".join(lines)


def test_apply_writes_with_backups(env):
    raws = setup(env)
    rb.rebuild(env, apply=True, imap_factory=lambda c, p: FakeIMAP(raws), password=PW, out=lambda s: None)
    d = json.loads(env.live_json.read_text(encoding="utf-8"))
    assert d["alert"] is False and not any(i["alert"] for i in d["items"])
    assert [r[1] for r in rows(env.stop_xlsx)[1:]] == ["contact@dead-shop.example"]
    assert rows(env.sent_xlsx)[1][5] == "отправлено; нерабочий"
    assert env.live_json.with_name("eaa_mail.json.bak").exists() and env.sent_xlsx.with_name(env.sent_xlsx.name + ".bak").exists()
    assert mw._read_json(env.state_json, {})["folders"]["INBOX"]["last_uid"] == 3                # курсор не тронут
    assert not env.state_json.with_name("mail_watch.lock").exists()


@pytest.mark.parametrize("cursor", [0, 3], ids=["no-cursor", "empty-mailbox"])
def test_refuses_when_zero_letters_read(env, cursor):
    setup(env)
    mw._write_json(env.state_json, {"live": [], "folders": {"INBOX": {"uidvalidity": 111, "last_uid": cursor}}})
    before = env.live_json.read_bytes()
    with pytest.raises(SystemExit, match="0 писем"):
        rb.rebuild(env, apply=True, imap_factory=lambda c, p: FakeIMAP([]), password=PW, out=lambda s: None)
    assert env.live_json.read_bytes() == before and not env.stop_xlsx.exists()
