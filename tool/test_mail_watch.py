"""Тесты mail_watch.py на подменённом ящике (сети нет, живого пароля нет). Запуск: python -m pytest tool/test_mail_watch.py"""
import json
import os
import re
import shutil
import sys
import time
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import mail_watch as mw  # noqa: E402

PW = "PWMARK-9f3a7c1d"
REAL_XLSX = Path("C:/!Projects/Office/workspace/eaa/круг1/Отправлено.xlsx")

HDR = ["Магазин", "Адрес", "Тема", "Файл", "Время", "Результат"]
SHOPS = [
    ("Dead Shop", "contact@dead-shop.example"),
    ("Mon Shop", "support@mon-shop.example"),
    ("Slip Shop", "bonjour@slip-shop.example"),
    ("Dmarc Shop", "support@dmarc-shop.example"),
    ("Live Shop", "info@live-shop.example"),
    ("Unsub Shop", "info@unsub-shop.example"),
    ("Quiet Shop", "hello@quiet-shop.example"),
]


def make_xlsx(path: Path, shops=SHOPS):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Отправлено"
    ws.append(HDR)
    for i, (n, a) in enumerate(shops):
        ws.append([n, a, f"Accessibility check of {n}", f"{n}.pdf", f"2026-11-10 00:0{i}:00", "отправлено"])
    wb.save(path)


def rows(path: Path):
    ws = openpyxl.load_workbook(path).active
    return [list(r) for r in ws.iter_rows(values_only=True)]


def mime(frm, subj, body, extra="", mid=None, ctype="text/plain; charset=utf-8"):
    mid = mid or f"<{abs(hash((frm, subj, body)))}@test>"
    return (f"From: {frm}\r\nTo: hello@eaa-audit.net\r\nSubject: {subj}\r\nDate: Tue, 10 Nov 2026 10:15:00 +0000\r\n"
            f"Message-ID: {mid}\r\n{extra}MIME-Version: 1.0\r\nContent-Type: {ctype}\r\n\r\n{body}\r\n").encode()


def dsn(addr, status="5.1.1", action="failed", diag="smtp; 550 5.1.1 No Such User Here", parsed=True):
    b = "BOUNDARY1"
    ds = (f"--{b}\r\nContent-Type: message/delivery-status\r\n\r\nReporting-MTA: dns; mail.privateemail.com\r\n\r\n"
          f"Final-Recipient: rfc822; {addr}\r\nAction: {action}\r\nStatus: {status}\r\nDiagnostic-Code: {diag}\r\n") if parsed else ""
    return (f"From: MAILER-DAEMON@mail.privateemail.com\r\nTo: hello@eaa-audit.net\r\n"
            f"Subject: Undelivered Mail Returned to Sender\r\nDate: Tue, 10 Nov 2026 09:00:00 +0000\r\n"
            f"Message-ID: <dsn1@test>\r\nMIME-Version: 1.0\r\n"
            f"Content-Type: multipart/report; report-type=delivery-status; boundary={b}\r\n\r\n"
            f"--{b}\r\nContent-Type: text/plain\r\n\r\nThis is the mail system. <{addr}>: host said: 550 No Such User Here\r\n"
            f"{ds}"
            f"--{b}\r\nContent-Type: text/rfc822-headers\r\n\r\nTo: {addr}\r\nSubject: Accessibility check\r\n"
            f"--{b}--\r\n").encode()


def dmarc():
    b = "BOUNDARY2"
    return (f"From: noreply-dmarc-support@google.com\r\nTo: hello@eaa-audit.net\r\n"
            f"Subject: Report domain: eaa-audit.net Submitter: google.com Report-ID: 123\r\n"
            f"Date: Tue, 10 Nov 2026 03:00:00 +0000\r\nMessage-ID: <dmarc1@test>\r\nMIME-Version: 1.0\r\n"
            f"Content-Type: multipart/mixed; boundary={b}\r\n\r\n"
            f"--{b}\r\nContent-Type: text/plain\r\n\r\nThis is an aggregate report.\r\n"
            f"--{b}\r\nContent-Type: application/zip; name=report.zip\r\nContent-Disposition: attachment; filename=report.zip\r\n"
            f"Content-Transfer-Encoding: base64\r\n\r\nUEsFBgAAAAAAAAAAAAAAAAAAAAAAAA==\r\n--{b}--\r\n").encode()


def six(shops=SHOPS):
    d = {s[0]: s[1] for s in shops}
    return [
        dsn(d["Dead Shop"]),
        mime(f"Mon Shop <{d['Mon Shop']}>", "Réponse automatique : Accessibility check", "Nous sommes absents.",
             extra="Auto-Submitted: auto-replied\r\n"),
        mime(f"Slip Shop <{d['Slip Shop']}>", "Re: Accessibility check of Slip Shop",
             "Bonjour,\r\n\r\nVotre demande a bien été transmise à notre service client. Nous revenons vers vous."),
        dmarc(),
        mime(f"Jean Dupont <jean.dupont@live-shop.example>", "Re: Accessibility check of Live Shop",
             "Bonjour, merci pour votre rapport. Pouvez-vous m'envoyer un devis pour la correction ? Cordialement, Jean",
             extra="In-Reply-To: <orig@x>\r\n"),
        mime(f"Unsub <{d['Unsub Shop']}>", "Re: Accessibility check of Unsub Shop",
             "Bonjour, merci de ne plus nous écrire.\r\n"),
    ]


class FakeIMAP:
    forbidden = ("store", "expunge", "delete", "copy", "move", "append", "create", "rename")

    def __init__(self, raws, validity=111, spam=None, spam_name="Junk", spam_flag=True, fail_uids=(), empty_uids=()):
        self.fail_uids, self.empty_uids = set(fail_uids), set(empty_uids)
        self.boxes = {"INBOX": raws}
        self.list_lines = [rb'(\HasNoChildren) "." "INBOX"']
        if spam is not None:
            self.boxes[spam_name] = spam
            fl = r"\HasNoChildren \Junk" if spam_flag else r"\HasNoChildren"
            self.list_lines.append(f'({fl}) "." "{spam_name}"'.encode())
        self.validity = validity
        self.selects = []          # [(папка, readonly)]
        self.cmds = []
        self.cur = None
        self.logged_out = False

    @property
    def readonly(self):
        return bool(self.selects) and all(ro is True for _, ro in self.selects)

    def __getattr__(self, name):
        if name in self.forbidden:
            raise AssertionError(f"запрещённая команда IMAP: {name}")
        raise AttributeError(name)

    def login(self, u, p):
        return "OK", []

    def list(self):
        return "OK", self.list_lines

    def select(self, box="INBOX", readonly=False):
        self.selects.append((box, readonly))
        if box.startswith('"'):
            box = box[1:-1]
        if box not in self.boxes:
            return "NO", []
        self.cur = box
        return "OK", [b"%d" % len(self.boxes[box])]

    def response(self, code):
        return code, [str(self.validity).encode()]

    def uid(self, cmd, *args):
        self.cmds.append((cmd, args))
        raws = self.boxes[self.cur]
        n = len(raws)
        if cmd == "SEARCH":
            if args[1] == "UID":
                lo = int(args[2].split(":")[0])
                ids = [u for u in range(1, n + 1) if u >= lo] or ([n] if n else [])
            else:
                ids = list(range(1, n + 1))
            return "OK", [b" ".join(b"%d" % i for i in ids)]
        if cmd == "FETCH":
            assert "BODY.PEEK" in args[1], "чтение без PEEK ставит флаг прочитано"
            if int(args[0]) in self.fail_uids:
                return "NO", [b"server busy"]
            if int(args[0]) in self.empty_uids:
                return "OK", [None]
            return "OK", [(b"1 (BODY[] {1}", raws[int(args[0]) - 1]), b")"]
        raise AssertionError(f"неожиданная команда {cmd}")

    def logout(self):
        self.logged_out = True


@pytest.fixture
def env(tmp_path, capsys):
    cfg = mw.Config(sent_xlsx=tmp_path / "Отправлено.xlsx", stop_xlsx=tmp_path / "Не_писать.xlsx",
                    live_json=tmp_path / "live" / "eaa_mail.json", state_json=tmp_path / "state" / "mail_watch.json",
                    log_file=tmp_path / "logs" / "mail_watch.log")
    make_xlsx(cfg.sent_xlsx)
    return cfg


def go(cfg, raws, validity=111, **kw):
    box = FakeIMAP(raws, validity, **kw)
    rc = mw.run(cfg, imap_factory=lambda c, p: box, password=PW)
    return rc, box


def test_six_kinds(env):
    cfg = env
    before = rows(cfg.sent_xlsx)
    rc, box = go(cfg, six())
    assert rc == 0
    after = rows(cfg.sent_xlsx)
    assert len(after) == len(before)
    for b, a in zip(before, after):                       # другие столбцы не тронуты
        assert b[:5] == a[:5]
    res = {r[0]: r[5] for r in after[1:]}
    assert "нерабочий" in res["Dead Shop"]
    assert "автоответ" in res["Mon Shop"]
    assert res["Slip Shop"] == "отправлено"               # «передано в отдел» без заголовка: живое «возможно автоответ»
    assert res["Dmarc Shop"] == "отправлено" and res["Quiet Shop"] == "отправлено"
    assert res["Live Shop"] == "отправлено"               # живой xlsx не трогает
    assert cfg.sent_xlsx.with_name(cfg.sent_xlsx.name + ".bak").exists()
    stop = rows(cfg.stop_xlsx)
    assert stop[0] == ["Магазин", "Адрес", "Причина", "Дата"]
    assert [r[1] for r in stop[1:]] == ["contact@dead-shop.example", "info@unsub-shop.example"]
    assert stop[1][0] == "Dead Shop" and stop[1][2].startswith("адрес не существует (5.1.1") and stop[2][0] == "Unsub Shop"
    live = json.loads(cfg.live_json.read_text(encoding="utf-8"))
    assert live["alert"] is True and len(live["items"]) == 2
    slip = [i for i in live["items"] if i.get("note") == "возможно автоответ"]
    assert len(slip) == 1 and slip[0]["alert"] is True and "slip-shop" in slip[0]["from"]
    it = [i for i in live["items"] if "jean.dupont" in i["from"]][0]
    assert it["alert"] is True and "jean.dupont@live-shop.example" in it["from"]
    assert it["store_domain"] == "live-shop.example" and it["subject"].startswith("Re: Accessibility")
    assert "devis" in it["text"] and len(it["text"]) <= 500 and it["message_id"] and "privateemail.com" in it["webmail_hint"]
    last = cfg.log_file.read_text(encoding="utf-8").strip().splitlines()[-1]
    assert "новых 6, недоставок 1, автоответов 1, отписок 1, служебных 1, живых 2" in last


def test_rerun_does_not_duplicate(env):
    cfg = env
    raws = six()
    go(cfg, raws)
    x1, l1, s1 = rows(cfg.sent_xlsx), cfg.live_json.read_text(encoding="utf-8"), rows(cfg.stop_xlsx)
    rc, box = go(cfg, raws)
    assert rc == 0
    assert rows(cfg.sent_xlsx) == x1 and rows(cfg.stop_xlsx) == s1
    assert cfg.live_json.read_text(encoding="utf-8").count('"message_id"') == 2
    assert "новых 0," in cfg.log_file.read_text(encoding="utf-8").strip().splitlines()[-1]
    assert box.cmds[0][1][:2] == (None, "UID") and box.cmds[0][1][2] == "7:*"   # по UID с прошлого запуска
    # состояние потеряно: те же письма не дают второго «живого» (дедуп по Message-ID) и не дублируют пометки
    cfg.state_json.unlink()
    go(cfg, raws)
    assert cfg.live_json.read_text(encoding="utf-8").count('"message_id"') == 2
    assert rows(cfg.sent_xlsx) == x1 and len(rows(cfg.stop_xlsx)) == len(s1)


def test_new_mail_after_first_run(env):
    cfg = env
    raws = six()
    go(cfg, raws)
    extra = mime("Boss <boss@quiet-shop.example>", "Re: check", "Intéressant, appelez-moi.")
    go(cfg, raws + [extra])
    items = json.loads(cfg.live_json.read_text(encoding="utf-8"))["items"]
    assert len(items) == 3 and items[2]["store_domain"] == "quiet-shop.example"


def test_readonly_and_peek(env):
    rc, box = go(env, six())
    assert box.readonly is True and box.logged_out
    assert {c for c, _ in box.cmds} <= {"SEARCH", "FETCH"}


def test_password_never_leaks(env, capsys):
    cfg = env

    def boom(c, p):
        raise OSError(f"login failed for {p}")
    rc = mw.run(cfg, imap_factory=boom, password=PW)
    assert rc == 1
    go(cfg, six())
    out = capsys.readouterr()
    blob = out.out + out.err
    for p in (cfg.log_file, cfg.state_json, cfg.live_json, cfg.sent_xlsx, cfg.stop_xlsx):
        blob += p.read_bytes().decode("utf-8", "ignore") if p.suffix != ".xlsx" else ""
    for p in (cfg.sent_xlsx, cfg.stop_xlsx):
        blob += " ".join(str(c) for r in rows(p) for c in r)
    assert PW not in blob and "PWMARK" not in blob
    assert "ошибка сети/IMAP" in cfg.log_file.read_text(encoding="utf-8")


def test_network_error_single_attempt(env):
    calls = []

    def boom(c, p):
        calls.append(1)
        raise TimeoutError("timed out")
    assert mw.run(env, imap_factory=boom, password=PW) == 1
    assert len(calls) == 1
    assert not env.live_json.exists()


def test_no_password(env, monkeypatch):
    monkeypatch.setattr(mw, "get_password", lambda: None)
    called = []
    assert mw.run(env, imap_factory=lambda c, p: called.append(1)) == 2
    assert not called
    assert "нет пароля, запусти Сохранить_пароль_почты.bat" in env.log_file.read_text(encoding="utf-8")


def test_doubtful_is_live(env):
    weird = mime("Somebody <x@unknown-domain.example>", "???", "ok")
    go(env, [weird])
    assert len(json.loads(env.live_json.read_text(encoding="utf-8"))["items"]) == 1


def test_newsletter_and_seo_are_silent(env):
    news = mime("News <news@promo.example>", "Our weekly offers", "Big sale.\r\nUnsubscribe here.",
                extra="List-Unsubscribe: <mailto:u@promo.example>\r\n")
    seo = mime("Bob <bob@seo-agency.example>", "Page 1 of Google for your site", "We do SEO and backlinks for you.")
    rc, _ = go(env, [news, seo])
    assert rc == 0 and not env.live_json.exists()
    assert "служебных 2, живых 0" in env.log_file.read_text(encoding="utf-8")
    assert not env.stop_xlsx.exists()


def test_unsub_footer_of_store_autoreply_is_not_unsub(env):
    m = mime("Mon <support@mon-shop.example>", "Re: check", "Thanks!\r\n" + "x" * 900 + "\r\nunsubscribe",
             extra="Auto-Submitted: auto-generated\r\n")
    go(env, [m])
    assert not env.stop_xlsx.exists()
    assert "автоответов 1" in env.log_file.read_text(encoding="utf-8")


def test_xlsx_locked_does_not_advance(env, monkeypatch):
    def fail(*a, **k):
        raise PermissionError("файл открыт в Excel")
    monkeypatch.setattr(mw, "_atomic_save", fail)
    rc, _ = go(env, six())
    assert rc == 3
    assert json.loads(env.state_json.read_text(encoding="utf-8")).get("last_uid", 0) == 0
    monkeypatch.undo()
    rc, _ = go(env, six())                                # повтор: пометки применены, живой не продублирован
    assert rc == 0 and len(json.loads(env.live_json.read_text(encoding="utf-8"))["items"]) == 2


def test_run_is_fast(env):
    t = time.time()
    go(env, six())
    assert time.time() - t < 5


def test_source_has_no_mutating_imap_calls():
    src = Path(mw.__file__).read_text(encoding="utf-8")
    for bad in ("store(", "expunge", "delete", "move"):
        assert bad not in src.lower(), bad
    assert "readonly=True" in src and "BODY.PEEK" in src
    assert PW not in src


@pytest.mark.skipif(not REAL_XLSX.exists(), reason="нет настоящего Отправлено.xlsx")
def test_on_copy_of_real_xlsx(env, tmp_path):
    """Прогон на копии настоящего Отправлено.xlsx: 6 видов с настоящими адресами из списка."""
    cfg = env
    shutil.copy2(REAL_XLSX, cfg.sent_xlsx)
    names = {r[1].lower(): r[0] for r in rows(cfg.sent_xlsx)[1:]}
    pick = {"dead": "contact@elogium.fr", "auto": "support@monbento.com", "slip": "bonjour@leslipfrancais.fr",
            "unsub": "bonjour@lepetitballon.com"}
    for a in pick.values():
        assert a in names
    raws = [dsn(pick["dead"]),
            mime(f"x <{pick['auto']}>", "Automatic reply", "We are away", extra="Auto-Submitted: auto-replied\r\n"),
            mime(f"x <{pick['slip']}>", "Re: Accessibility check", "Votre message a bien été transmis à notre équipe."),
            dmarc(),
            mime("Marie <marie@ldlc.com>", "Re: Accessibility check of LDLC", "Bonjour, pouvez-vous nous rappeler ?"),
            mime(f"x <{pick['unsub']}>", "Re: Accessibility check", "Ne plus recevoir vos messages, merci.")]
    before = rows(cfg.sent_xlsx)
    rc, _ = go(cfg, raws)
    assert rc == 0
    after = rows(cfg.sent_xlsx)
    assert len(before) == len(after)
    diff = {a[1]: a[5] for b, a in zip(before, after) if b != a}
    assert set(diff) == {pick["dead"], pick["auto"]}
    assert "нерабочий" in diff[pick["dead"]] and "автоответ" in diff[pick["auto"]]
    assert [r[1] for r in rows(cfg.stop_xlsx)[1:]] == [pick["dead"], pick["unsub"]]
    items = json.loads(cfg.live_json.read_text(encoding="utf-8"))["items"]
    assert {i["store_domain"] for i in items} == {"ldlc.com", "leslipfrancais.fr"}


def test_dead_address_goes_to_do_not_write_once(env):
    go(env, [dsn("contact@dead-shop.example")])
    go(env, [dsn("contact@dead-shop.example"), dsn("contact@dead-shop.example")])   # тот же адрес ещё раз
    stop = rows(env.stop_xlsx)
    assert [r[1] for r in stop[1:]] == ["contact@dead-shop.example"] and stop[1][2].startswith("адрес не существует")
    assert "нерабочий" in rows(env.sent_xlsx)[1][5]
    assert rows(env.sent_xlsx)[1][5].count("нерабочий") == 1


def test_dead_address_appended_to_existing_stop_list(env):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Не писать"
    ws.append(["Магазин", "Адрес", "Причина", "Дата"])
    ws.append(["Old", "old@old.example", "недоставка", "2026-11-09"])
    wb.save(env.stop_xlsx)
    go(env, [dsn("contact@dead-shop.example")])
    stop = rows(env.stop_xlsx)
    assert [r[1] for r in stop[1:]] == ["old@old.example", "contact@dead-shop.example"]
    assert env.stop_xlsx.with_name(env.stop_xlsx.name + ".bak").exists()


@pytest.mark.parametrize("name,flag", [("Junk", True), ("Spam", False), ("INBOX.Spam", False)])
def test_live_reply_in_spam_alerts(env, name, flag):
    reply = mime("Jean <jean@live-shop.example>", "Re: Accessibility check", "Bonjour, envoyez le devis.",
                 mid="<spam-live@test>")
    rc, box = go(env, [], spam=[reply], spam_name=name, spam_flag=flag)
    assert rc == 0
    d = json.loads(env.live_json.read_text(encoding="utf-8"))
    assert d["alert"] is True and len(d["items"]) == 1
    it = d["items"][0]
    assert it["alert"] is True and it["folder"] == name and "спам" in it["webmail_hint"]
    assert box.readonly is True and [f for f, _ in box.selects] == ["INBOX", name]


def test_real_spam_in_spam_folder_is_silent(env):
    junk = [mime("Bob <bob@casino.example>", "You won", "Click here"),
            mime("Seo <s@seo.example>", "Rank", "We do SEO")]
    rc, _ = go(env, [], spam=junk)
    assert rc == 0 and not env.live_json.exists()
    assert "служебных 2, живых 0" in env.log_file.read_text(encoding="utf-8")


def test_uid_tracked_per_folder(env):
    inbox = [mime("A <a@quiet-shop.example>", "Re: one", "un", mid="<i1@t>")]
    spam = [mime("B <b@live-shop.example>", "Re: two", "deux", mid="<s1@t>"),
            mime("C <c@live-shop.example>", "Re: three", "trois", mid="<s2@t>")]
    go(env, inbox, spam=spam)
    st = json.loads(env.state_json.read_text(encoding="utf-8"))["folders"]
    assert st["INBOX"]["last_uid"] == 1 and st["Junk"]["last_uid"] == 2
    rc, box = go(env, inbox + [mime("D <d@quiet-shop.example>", "Re: four", "quatre", mid="<i2@t>")], spam=spam)
    assert [a[2] for c, a in box.cmds if c == "SEARCH"] == ["2:*", "3:*"]     # INBOX с UID 2, спам с UID 3
    assert len(json.loads(env.live_json.read_text(encoding="utf-8"))["items"]) == 4


def test_spam_folder_unreadable_is_not_fatal(env):
    box = FakeIMAP([mime("A <a@quiet-shop.example>", "Re: one", "un")], spam=[])
    box.list_lines.append(rb'(\HasNoChildren \Junk) "." "Gone"')
    del box.boxes["Junk"]
    rc = mw.run(env, imap_factory=lambda c, p: box, password=PW)
    assert rc == 1                                        # ошибка видна, но INBOX обработан
    assert len(json.loads(env.live_json.read_text(encoding="utf-8"))["items"]) == 1


def test_old_state_format_migrates(env):
    env.state_json.parent.mkdir(parents=True)
    env.state_json.write_text(json.dumps({"last_uid": 1, "uidvalidity": 111, "live": []}), encoding="utf-8")
    rc, box = go(env, six()[:1] + six()[4:5])
    assert box.cmds[0][1][2] == "2:*"
    assert json.loads(env.state_json.read_text(encoding="utf-8"))["folders"]["INBOX"]["last_uid"] == 2


# ------------------------------------------------------------ круг по reviewer: пункты 1-7 и дополнительные

def live_ids(cfg):
    return [i["message_id"] for i in json.loads(cfg.live_json.read_text(encoding="utf-8"))["items"]]


def live_items(cfg):
    return json.loads(cfg.live_json.read_text(encoding="utf-8"))["items"]


def state_of(cfg):
    return json.loads(cfg.state_json.read_text(encoding="utf-8"))


def three_live():
    return [mime(f"P{i} <p{i}@live-shop.example>", f"Re: letter {i}", f"réponse {i}", mid=f"<m{i}@t>") for i in (1, 2, 3)]


@pytest.mark.parametrize("kw", [dict(fail_uids=[2]), dict(empty_uids=[2])])
def test_p1_fetch_failure_holds_cursor(env, kw):
    rc, box = go(env, three_live(), **kw)
    assert rc == 1
    assert live_ids(env) == ["<m1@t>"]                       # письмо до сбоя обработано, после — не пропущено
    assert state_of(env)["folders"]["INBOX"]["last_uid"] == 1
    rc, _ = go(env, three_live())                            # сбой прошёл: 2 и 3 подобраны, дублей нет
    assert rc == 0 and live_ids(env) == ["<m1@t>", "<m2@t>", "<m3@t>"]
    assert state_of(env)["folders"]["INBOX"]["last_uid"] == 3


def test_p1_first_message_fails_cursor_stays_zero(env):
    rc, _ = go(env, three_live(), fail_uids=[1])
    assert rc == 1 and not env.live_json.exists()
    assert state_of(env)["folders"]["INBOX"].get("last_uid", 0) == 0


def test_p2_json_write_failure_holds_cursor(env, monkeypatch):
    real = mw._write_json

    def flaky(p, data):
        if p.name == "eaa_mail.json":
            raise OSError("диск занят")
        return real(p, data)
    monkeypatch.setattr(mw, "_write_json", flaky)
    rc, _ = go(env, three_live())
    assert rc == 3
    st = state_of(env)
    assert st["folders"]["INBOX"].get("last_uid", 0) == 0 and st["live"] == []   # живое не считается виденным
    monkeypatch.undo()
    rc, _ = go(env, three_live())
    assert rc == 0 and live_ids(env) == ["<m1@t>", "<m2@t>", "<m3@t>"]
    assert state_of(env)["folders"]["INBOX"]["last_uid"] == 3


def test_p3_corrupt_live_json_is_not_overwritten(env):
    env.live_json.parent.mkdir(parents=True)
    env.live_json.write_text("{ тревоги: [битый json", encoding="utf-8")
    rc, _ = go(env, three_live())
    assert rc == 3
    assert env.live_json.read_text(encoding="utf-8") == "{ тревоги: [битый json"
    assert env.live_json.with_name("eaa_mail.json.corrupt").read_text(encoding="utf-8") == "{ тревоги: [битый json"
    assert state_of(env)["folders"]["INBOX"].get("last_uid", 0) == 0
    assert "нечитаем" in env.log_file.read_text(encoding="utf-8")


def test_p3_existing_alerts_are_kept(env):
    go(env, three_live()[:1])
    go(env, three_live())
    assert live_ids(env) == ["<m1@t>", "<m2@t>", "<m3@t>"]


def test_p4_validity_change_with_empty_mailbox_resets_cursor(env):
    go(env, three_live(), validity=111)
    assert state_of(env)["folders"]["INBOX"]["last_uid"] == 3
    rc, _ = go(env, [], validity=222)
    f = state_of(env)["folders"]["INBOX"]
    assert rc == 0 and f["last_uid"] == 0 and f["uidvalidity"] == 222


def test_p4_validity_change_rereads_without_duplicates(env):
    go(env, three_live(), validity=111)
    rc, box = go(env, three_live(), validity=222)             # ящик пересоздан с теми же письмами
    assert rc == 0 and box.cmds[0][1][1] == "SINCE"
    assert live_ids(env) == ["<m1@t>", "<m2@t>", "<m3@t>"]
    assert state_of(env)["folders"]["INBOX"] == {"last_uid": 3, "uidvalidity": 222}


def test_p5_outlook_header_does_not_hide_human_reply(env):
    m = mime("Sean <sean@live-shop.example>", "RE: Accessibility check", "Hi, can you call me tomorrow?",
             extra="X-Auto-Response-Suppress: DR, OOF, AutoReply\r\nIn-Reply-To: <orig@x>\r\n")
    go(env, [m])
    items = live_items(env)
    assert len(items) == 1 and items[0]["alert"] is True and "note" not in items[0]


@pytest.mark.parametrize("status,action", [("4.4.1", "delayed"), ("4.2.2", "failed"), ("5.2.2", "failed"),
                                           ("5.7.1", "failed"), ("5.1.1", "delayed")])
def test_p6_soft_dsn_is_not_do_not_write(env, status, action):
    rc, _ = go(env, [dsn("contact@dead-shop.example", status=status, action=action, diag="smtp; mailbox full")])
    assert rc == 0
    cell = rows(env.sent_xlsx)[1][5]
    assert "недоставка временно/отказ" in cell and "нерабочий" not in cell
    assert not env.stop_xlsx.exists()


def test_p6_hard_dsn_reason_is_real(env):
    go(env, [dsn("contact@dead-shop.example", status="5.1.2", diag="smtp; 550 5.1.2 Host not found")])
    stop = rows(env.stop_xlsx)
    assert "нерабочий" in rows(env.sent_xlsx)[1][5]
    assert "5.1.2" in stop[1][2] and "Host not found" in stop[1][2]


def test_p6_unparsed_dsn_alerts_and_marks_nothing(env):
    before = rows(env.sent_xlsx)
    go(env, [dsn("contact@dead-shop.example", parsed=False)])
    assert rows(env.sent_xlsx) == before and not env.stop_xlsx.exists()
    items = live_items(env)
    assert len(items) == 1 and items[0]["alert"] is True and items[0]["note"] == "недоставка не разобрана"


def test_p7_phrase_only_before_quote(env):
    body = ("Merci, appelez-moi au 06 00 00 00 00.\r\n\r\nLe lun. 9 nov. 2026 à 23:56, Hello <hello@eaa-audit.net> a écrit :\r\n"
            "> Bonjour,\r\n> Votre demande a bien été transmise à notre service client.\r\n")
    go(env, [mime("Jean <jean@live-shop.example>", "Re: Accessibility check", body)])
    items = live_items(env)
    assert len(items) == 1 and "note" not in items[0]
    body_en = "Sounds good, send the quote.\r\n\r\nOn Mon, Nov 9, 2026 at 23:56 Hello wrote:\r\n> we received your request\r\n"
    go(env, [mime("Jean <jean@live-shop.example>", "Re: Accessibility check", body)] +
       [mime("A <a@x1.example>", "Re: s", body_en, mid="<en@t>")])
    assert "note" not in live_items(env)[-1] and len(live_items(env)) == 2


def test_p7_text_only_autoreply_is_live_with_note(env):
    m = mime("Slip <bonjour@slip-shop.example>", "Re: Accessibility check", "Votre demande a bien été transmise à notre équipe.")
    go(env, [m])
    it = live_items(env)[0]
    assert it["note"] == "возможно автоответ" and it["alert"] is True
    assert rows(env.sent_xlsx)[3][5] == "отправлено"


def test_p7_silent_categories_are_logged_with_sender_and_subject(env):
    go(env, six())
    log_text = env.log_file.read_text(encoding="utf-8")
    assert "автоответ [INBOX UID 2] от support@mon-shop.example: «Réponse automatique : Accessibility check»" in log_text
    assert "отписка [INBOX UID 6] от info@unsub-shop.example" in log_text
    assert "служебное [INBOX UID 4] от noreply-dmarc-support@google.com: «Report domain: eaa-audit.net" in log_text


def test_spam_service_is_logged(env):
    go(env, [], spam=[mime("Bob <bob@casino.example>", "You won", "Click here")])
    assert "служебное [Junk UID 1] от bob@casino.example: «You won»" in env.log_file.read_text(encoding="utf-8")


def test_generic_domain_matches_only_full_address(env):
    make_xlsx(env.sent_xlsx, SHOPS + [("Gmail Shop", "shop@gmail.com")])
    auto = mime("Other <other@gmail.com>", "Automatic reply", "away", extra="Auto-Submitted: auto-replied\r\n")
    unsub = mime("Other2 <other2@gmail.com>", "Re: x", "ne plus recevoir vos messages", mid="<g2@t>")
    own = mime("Shop <shop@gmail.com>", "Automatic reply", "away", extra="Auto-Submitted: auto-replied\r\n", mid="<g3@t>")
    go(env, [auto, unsub, own])
    res = {r[1]: r[5] for r in rows(env.sent_xlsx)[1:]}
    assert res["shop@gmail.com"].count("автоответ") == 1       # только от самого shop@gmail.com
    assert not env.stop_xlsx.exists()                             # чужой gmail не «отписка магазина»
    items = live_items(env)
    assert len(items) == 1 and items[0]["store_domain"] == "gmail.com" and "other2@gmail.com" in items[0]["from"]


def test_spam_folder_name_with_space_is_quoted(env):
    reply = mime("Jean <jean@live-shop.example>", "Re: x", "bonjour, devis ?", mid="<sp@t>")
    rc, box = go(env, [], spam=[reply], spam_name="Junk Mail")
    assert rc == 0 and box.selects[1][0] == '"Junk Mail"'
    assert len(live_ids(env)) == 1


def test_lock_blocks_parallel_run(env):
    lock = env.state_json.with_name("mail_watch.lock")
    lock.parent.mkdir(parents=True)
    lock.write_text("123")
    called = []
    assert mw.run(env, imap_factory=lambda c, p: called.append(1), password=PW) == 4
    assert not called and lock.exists()
    old = time.time() - 3600
    os.utime(lock, (old, old))                                    # замок от упавшего запуска
    rc, _ = go(env, three_live())
    assert rc == 0 and not lock.exists()
    rc, _ = go(env, three_live())                                 # замок снимается после каждого запуска
    assert rc == 0


def test_lock_released_after_network_error(env):
    def boom(c, p):
        raise TimeoutError("x")
    assert mw.run(env, imap_factory=boom, password=PW) == 1
    assert not env.state_json.with_name("mail_watch.lock").exists()


def test_poison_message_is_live_not_lost(env, monkeypatch):
    real = mw.classify

    def bad(msg, *a, **k):
        if "poison" in str(msg.get("Subject")):
            raise RuntimeError("boom")
        return real(msg, *a, **k)
    monkeypatch.setattr(mw, "classify", bad)
    rc, _ = go(env, [mime("X <x@live-shop.example>", "poison", "текст"), three_live()[0]])
    assert rc == 0 and len(live_ids(env)) == 2
    assert live_items(env)[0]["note"] == "ошибка разбора"


def test_check_password(capsys):
    cfg = mw.Config()

    class OK:
        def logout(self):
            pass
    assert mw.check_password(cfg, imap_factory=lambda c, p: OK(), password=PW) == 0
    assert "Пароль верный" in capsys.readouterr().out

    def bad(c, p):
        raise mw.imaplib.IMAP4.error(f"[AUTHENTICATIONFAILED] {p}")
    assert mw.check_password(cfg, imap_factory=bad, password=PW) == 1
    out = capsys.readouterr()
    assert "НЕ подошёл" in out.out and PW not in out.out + out.err

    def down(c, p):
        raise TimeoutError(p)
    assert mw.check_password(cfg, imap_factory=down, password=PW) == 3
    out = capsys.readouterr()
    assert PW not in out.out + out.err and "не проверен" in out.out


def test_clipboard_cleared_even_if_save_fails(monkeypatch):
    calls = []

    class R:
        stdout = PW + "\n"

    def fake_run(cmd, **k):
        calls.append(cmd[-1])
        return R()
    monkeypatch.setattr(mw.subprocess, "run", fake_run)
    monkeypatch.setattr(mw, "_dpapi", lambda d, p: (_ for _ in ()).throw(OSError("dpapi")))
    with pytest.raises(OSError):
        mw.save_password()
    assert any("Set-Clipboard" in c for c in calls)
