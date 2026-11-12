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


def make_xlsx(path: Path, shops=SHOPS, subjects=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Отправлено"
    ws.append(HDR)
    for i, (n, a) in enumerate(shops):
        subj = (subjects or {}).get(n, f"Accessibility check of {n}")
        ws.append([n, a, subj, f"{n}.pdf", f"2026-11-10 00:0{i}:00", "отправлено"])
    wb.save(path)


def rows(path: Path):
    ws = openpyxl.load_workbook(path).active
    return [list(r) for r in ws.iter_rows(values_only=True)]


def mime(frm, subj, body, extra="", mid=None, ctype="text/plain; charset=utf-8"):
    mid = mid or f"<{abs(hash((frm, subj, body)))}@test>"
    return (f"From: {frm}\r\nTo: hello@eaa-audit.net\r\nSubject: {subj}\r\nDate: Tue, 10 Nov 2026 10:15:00 +0000\r\n"
            f"Message-ID: {mid}\r\n{extra}MIME-Version: 1.0\r\nContent-Type: {ctype}\r\n\r\n{body}\r\n").encode()


def dsn(addr, status="5.1.1", action="failed", diag="smtp; 550 5.1.1 No Such User Here", parsed=True,
        text="host said: 550 No Such User Here"):
    b = "BOUNDARY1"
    ds = (f"--{b}\r\nContent-Type: message/delivery-status\r\n\r\nReporting-MTA: dns; mail.privateemail.com\r\n\r\n"
          f"Final-Recipient: rfc822; {addr}\r\nAction: {action}\r\nStatus: {status}\r\nDiagnostic-Code: {diag}\r\n") if parsed else ""
    return (f"From: MAILER-DAEMON@mail.privateemail.com\r\nTo: hello@eaa-audit.net\r\n"
            f"Subject: Undelivered Mail Returned to Sender\r\nDate: Tue, 10 Nov 2026 09:00:00 +0000\r\n"
            f"Message-ID: <dsn1@test>\r\nMIME-Version: 1.0\r\n"
            f"Content-Type: multipart/report; report-type=delivery-status; boundary={b}\r\n\r\n"
            f"--{b}\r\nContent-Type: text/plain\r\n\r\nThis is the mail system. <{addr}>: {text}\r\n"
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

    def __init__(self, raws, validity=111, spam=None, spam_name="Junk", spam_flag=True, fail_uids=(), empty_uids=(), on_fetch=None):
        self.fail_uids, self.empty_uids = set(fail_uids), set(empty_uids)
        self.on_fetch = on_fetch
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
            if self.on_fetch:
                self.on_fetch(int(args[0]))
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
    wb = openpyxl.load_workbook(cfg.sent_xlsx)             # живой прогон 12.11 уже поставил метки в настоящий файл: в копии их снимаем
    for row in wb.active.iter_rows(min_row=2):
        if str(row[5].value or "").startswith("отправлено"):
            row[5].value = "отправлено"
    wb.save(cfg.sent_xlsx)
    names = {r[1].lower(): r[0] for r in rows(cfg.sent_xlsx)[1:]}
    by_dom = {}                                            # по одному адресу с каждого домена магазина, из самого файла
    for a in names:
        if mw._domain(a) not in mw.GENERIC_DOMAINS:
            by_dom.setdefault(mw._domain(a), a)
    assert len(by_dom) >= 5
    dead, auto, slip, unsub, person_dom = list(by_dom.values())[:4] + [list(by_dom)[4]]
    pick = {"dead": dead, "auto": auto, "slip": slip, "unsub": unsub}
    raws = [dsn(pick["dead"]),
            mime(f"x <{pick['auto']}>", "Automatic reply", "We are away", extra="Auto-Submitted: auto-replied\r\n"),
            mime(f"x <{pick['slip']}>", "Re: Accessibility check", "Votre message a bien été transmis à notre équipe."),
            dmarc(),
            mime(f"Marie <marie@{person_dom}>", "Re: Accessibility check", "Bonjour, pouvez-vous nous rappeler ?"),
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
    assert {i["store_domain"] for i in items} == {person_dom, mw._domain(slip)}


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


# ------------------------------------------------------------ хвосты после ПРИНЯТО reviewer (5 пунктов)

def without_message_id(raw: bytes) -> bytes:
    return re.sub(rb"Message-ID:[^\r\n]*\r\n", b"", raw)


def test_t1_dedup_key_never_empty_without_message_id(env, monkeypatch):
    live = without_message_id(mime("Jean <jean@live-shop.example>", "Re: devis", "bonjour, un devis ?"))
    assert b"Message-ID" not in live
    auto = mime("Mon <support@mon-shop.example>", "Automatic reply", "away", extra="Auto-Submitted: auto-replied\r\n")
    real = mw._atomic_save

    def fail_xlsx(*a, **k):
        raise PermissionError("xlsx открыт в Excel")
    monkeypatch.setattr(mw, "_atomic_save", fail_xlsx)
    rc, _ = go(env, [live, auto])                      # json записан, xlsx нет -> курсор стоит
    assert rc == 3 and len(live_items(env)) == 1
    assert live_items(env)[0]["message_id"] == "INBOX:1"
    monkeypatch.setattr(mw, "_atomic_save", real)
    rc, _ = go(env, [live, auto])                      # повтор: то же письмо второй раз
    assert rc == 0 and len(live_items(env)) == 1
    assert "автоответ" in rows(env.sent_xlsx)[2][5]


def test_t2_corrupt_copy_made_once_over_24_runs(env):
    env.live_json.parent.mkdir(parents=True)
    env.live_json.write_text("битый{", encoding="utf-8")
    for i in range(24):
        rc, _ = go(env, [mime("P <p@live-shop.example>", "Re: x", "bonjour", mid=f"<c{i}@t>")])
        assert rc == 3
    assert sorted(f.name for f in env.live_json.parent.iterdir()) == ["eaa_mail.json", "eaa_mail.json.corrupt"]


def test_t3_dsn_alias_recipient_found_in_text(env):
    raw = dsn("contact@dead-shop.example")
    raw = raw.replace(b"Final-Recipient: rfc822; contact@dead-shop.example", b"Final-Recipient: rfc822; alias@forward.example")
    assert b"alias@forward.example" in raw
    rc, _ = go(env, [raw])
    assert rc == 0 and "нерабочий" in rows(env.sent_xlsx)[1][5]
    assert [r[1] for r in rows(env.stop_xlsx)[1:]] == ["contact@dead-shop.example"]


def test_t4_check_password_network_vs_refusal(capsys):
    cfg = mw.Config()
    for exc in (mw.imaplib.IMAP4.abort("socket error: EOF"), TimeoutError("t"), OSError("dns"),
                mw.imaplib.IMAP4.error("[UNAVAILABLE] try later")):
        def f(c, p, exc=exc):
            raise exc
        assert mw.check_password(cfg, imap_factory=f, password=PW) == 3
        out = capsys.readouterr().out
        assert "Нет связи с почтой, попробуйте позже" in out and "НЕ подошёл" not in out

    def refuse(c, p):
        raise mw.imaplib.IMAP4.error("[AUTHENTICATIONFAILED] Authentication failed.")
    assert mw.check_password(cfg, imap_factory=refuse, password=PW) == 1
    out = capsys.readouterr().out
    assert "НЕ подошёл" in out and "Нет связи" not in out and PW not in out


def test_t5_lock_refreshed_on_each_message(env, monkeypatch):
    clock = [time.time()]
    monkeypatch.setattr(mw.time, "time", lambda: clock[0])
    other = []

    def on_fetch(uid):
        clock[0] += 400                                   # каждое письмо «читается» 400 с; всего 1200 с > 600
        if uid == 3:                                      # второй запуск пытается перехватить замок
            called = []
            other.append((mw.run(env, imap_factory=lambda c, p: called.append(1), password=PW), called))
    rc, _ = go(env, three_live(), on_fetch=on_fetch)
    assert rc == 0 and other == [(4, [])]                 # второй получил код 4 и ящик не открывал
    assert live_ids(env) == ["<m1@t>", "<m2@t>", "<m3@t>"]


# ---------------------------------------------------------------- 12.11, по живому запуску tester: Д1–Д3

def dmarc_from(frm, subj="Report Domain: eaa-audit.net Submitter: x.example Report-ID: 1", text="Find attached the DMARC Aggregate Report.",
               attach="x.example!eaa-audit.net!1!2.xml.gz", mid="<dm@test>"):
    b = "BOUNDARY3"
    att = (f"--{b}\r\nContent-Type: application/gzip; name=\"{attach}\"\r\nContent-Disposition: attachment; filename=\"{attach}\"\r\n"
           f"Content-Transfer-Encoding: base64\r\n\r\nH4sIAAAAAAAAAAMAAAAAAAAAAAA=\r\n") if attach else ""
    return (f"From: {frm}\r\nTo: hello@eaa-audit.net\r\nSubject: {subj}\r\nDate: Wed, 11 Nov 2026 03:00:00 +0000\r\n"
            f"Message-ID: {mid}\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary={b}\r\n\r\n"
            f"--{b}\r\nContent-Type: text/plain\r\n\r\n{text}\r\n{att}--{b}--\r\n").encode()


def log_text(cfg):
    return cfg.log_file.read_text(encoding="utf-8")


def test_d1_dmarc_from_mailer_daemon_is_service_not_alert(env):
    raws = [dmarc_from("MAILER-DAEMON@esa2.mx-relay.example", mid="<a@t>"),
            dmarc_from("MAILER-DAEMON@tech-group.example", subj="Report Domain: eaa-audit.net Submitter: mail.tech-group.example Report-ID: 9", mid="<b@t>")]
    rc, _ = go(env, raws)
    assert rc == 0
    assert not env.live_json.exists()                                   # ни одной тревоги
    text = log_text(env)
    assert text.count("служебное [INBOX UID") == 2 and "(DMARC)" in text and "недоставка не разобрана" not in text
    assert "недоставок 0" in text.splitlines()[-1] and "служебных 2, живых 0" in text.splitlines()[-1]
    assert not env.stop_xlsx.exists() and rows(env.sent_xlsx)[1][5] == "отправлено"


@pytest.mark.parametrize("kw", [
    dict(subj="Daily summary", attach=None),                                       # признак только в тексте
    dict(subj="Report for eaa-audit.net", text="Nothing here"),                    # только вложение .xml.gz (+ слово report в теме)
    dict(subj="rapport agrégé eaa-audit.net", text="Voir pièce jointe", attach=None),               # только тема
], ids=["text", "attachment", "subject"])
def test_d1_each_dmarc_sign_is_enough_before_daemon_check(env, kw):
    go(env, [dmarc_from("MAILER-DAEMON@mail.example.org", **kw)])
    assert not env.live_json.exists() and "служебное [INBOX UID 1] от mailer-daemon@mail.example.org" in log_text(env)


def test_d1_real_delivery_report_stays_dsn(env):
    go(env, [dsn("contact@dead-shop.example")])
    assert "нерабочий" in rows(env.sent_xlsx)[1][5] and "служебное" not in log_text(env)


def promo(frm="Wine Shop <bonjour@slip-shop.example>", extra="List-Unsubscribe: <mailto:unsub@x.example>\r\n",
          subj="Vous risquez d'y prendre goût 😉"):
    return mime(frm, subj, "https://www.slip-shop.example/pages/15-ans\r\nhttps://www.slip-shop.example/products/abonnement", extra=extra)


@pytest.mark.parametrize("extra", ["List-Unsubscribe: <mailto:unsub@x.example>\r\n", "Precedence: bulk\r\n"],
                         ids=["list-unsubscribe", "precedence-bulk"])
def test_d2_store_newsletter_is_service_not_live(env, extra):
    go(env, [promo(extra=extra)])
    assert not env.live_json.exists()
    assert "служебное [INBOX UID 1] от bonjour@slip-shop.example" in log_text(env) and "служебных 1, живых 0" in log_text(env)


@pytest.mark.parametrize("extra,subj", [
    ("List-Unsubscribe: <mailto:u@x.example>\r\nIn-Reply-To: <our@eaa-audit.net>\r\n", "Accessibility check of Slip Shop"),
    ("List-Unsubscribe: <mailto:u@x.example>\r\nReferences: <our@eaa-audit.net>\r\n", "Accessibility check of Slip Shop"),
    ("Precedence: bulk\r\n", "Re: Accessibility check of Slip Shop"),
    ("Precedence: bulk\r\n", "TR: Accessibility check of Slip Shop"),
], ids=["in-reply-to", "references", "Re", "TR"])
def test_d2_reply_to_us_from_store_stays_live_even_with_bulk_headers(env, extra, subj):
    go(env, [promo(extra=extra, subj=subj)])
    assert len(live_items(env)) == 1 and live_items(env)[0]["alert"] is True


def test_d2_human_letter_without_bulk_signs_is_live(env):
    m = mime("Marie <marie@slip-shop.example>", "Accessibilité de Slip Shop : 19 éléments à corriger",
             "Bonjour, merci de penser à nous. Je fais suivre votre dossier auprès du service concerné qui s’occupera de vous recontacter.")
    go(env, [m])
    assert len(live_items(env)) == 1


def test_d2_non_store_bulk_is_service_even_with_in_reply_to(env):
    go(env, [mime("News <news@other.example>", "Re: weekly", "hello",
                  extra="List-Unsubscribe: <mailto:u@o.example>\r\nIn-Reply-To: <x@y>\r\n")])
    assert not env.live_json.exists()


def hard_case(env, status="5.0.0", diag="smtp; 550 No Such User Here", action="failed", **kw):
    rc, _ = go(env, [dsn("contact@dead-shop.example", status=status, action=action, diag=diag, **kw)])
    assert rc == 0
    return rows(env.sent_xlsx)[1][5], (rows(env.stop_xlsx) if env.stop_xlsx.exists() else [])


@pytest.mark.parametrize("code", ["550", "551", "553"])
def test_d3_status_500_with_smtp_code_is_permanent(env, code):
    cell, stop = hard_case(env, diag=f"smtp; {code} No Such User Here")
    assert "нерабочий" in cell and "недоставка временно/отказ" not in cell
    assert [r[1] for r in stop[1:]] == ["contact@dead-shop.example"] and code in stop[1][2] and "5.0.0" in stop[1][2]


def test_d3_office365_recipient_rejected_is_permanent(env):
    cell, stop = hard_case(env, diag="smtp; 550 5.4.1 Recipient address rejected: Access denied. AS(201806281)")
    assert "нерабочий" in cell and len(stop) == 2


def test_d3_code_only_in_message_text_is_permanent(env):
    cell, stop = hard_case(env, diag="", text="host said: 550 No Such User Here")
    assert "нерабочий" in cell and len(stop) == 2 and "550" in stop[1][2]


@pytest.mark.parametrize("status,action,diag,text", [
    ("5.0.0", "failed", "smtp; 550 Message rejected as spam", "x"),               # код есть, но это блокировка
    ("5.0.0", "failed", "smtp; 550 blocked by reputation list", "x"),
    ("5.0.0", "failed", "smtp; mailbox unavailable", "Delivery problem, no code here"),   # кода нет
    ("5.0.0", "failed", "smtp; 554 transaction failed", "x"),                       # другой код
    ("5.0.0", "delayed", "smtp; 550 No Such User Here", "x"),                       # не failed
    ("5.2.2", "failed", "smtp; 550 5.2.2 mailbox full", "x"),
    ("5.7.1", "failed", "smtp; 550 5.7.1 relay denied", "x"),
    ("4.4.1", "failed", "smtp; 550 No Such User Here", "x"),
], ids=["spam", "reputation", "no-code", "554", "delayed", "5.2.2", "5.7.1", "4.4.1"])
def test_d3_not_permanent_stays_soft(env, status, action, diag, text):
    cell, stop = hard_case(env, status=status, action=action, diag=diag, text=text)
    assert "недоставка временно/отказ" in cell and "нерабочий" not in cell and stop == []


def test_replay_live_run_12_11(env):
    """Перепрогон классификации по виду писем живого запуска 12.11 (15 писем; сети нет, письма синтетические по отчёту tester)."""
    cfg = env
    shops = [("Pharmacy Shop", "info@pharmacy-shop.example"), ("DIY Shop", "sales@diy-shop.example"), ("Slip Shop", "bonjour@slip-shop.example"),
             ("Tech Shop", "support@tech-shop.example"), ("Wine Shop", "bonjour@wine-shop.example"), ("Lunch Shop", "support@lunch-shop.example")]
    make_xlsx(cfg.sent_xlsx, shops)
    subj = "Accessibilité de Slip Shop : 19 éléments à corriger"
    auto = "Auto-Submitted: auto-replied\r\n"
    raws = [
        dsn("info@pharmacy-shop.example", status="5.0.0", diag="smtp; 550 No Such User Here"),
        dsn("sales@diy-shop.example", status="5.0.0", diag="smtp; 550 5.4.1 Recipient address rejected: Access denied"),
        mime("Slip Shop <bonjour@slip-shop.example>", subj, "Bonjour, accusé de réception.", extra=auto, mid="<s1@t>"),
        mime("Tech Shop <support@tech-shop.example>", "Bonne réception de votre message", "Ceci est un message automatique.", extra=auto, mid="<l1@t>"),
        mime("Wine Shop <bonjour@wine-shop.example>", "Votre demande a bien été reçue", "Merci.", extra=auto, mid="<p1@t>"),
        mime("Lunch Shop <support@lunch-shop.example>", "Votre demande a bien été reçue chez Lunch Shop", "Merci.", extra=auto, mid="<m1@t>"),
        mime("Slip Shop <bonjour@slip-shop.example>", subj,
             "Bonjour, Merci de penser à nous, c’est vraiment sympa ! Je fais suivre votre dossier auprès du service concerné "
             "qui s’occupera de vous recontacter.", mid="<s2@t>"),                                   # единственный живой ответ
        dmarc_from("noreply-dmarc-support@google.com", mid="<d1@t>"),
        dmarc_from("no-reply@eu-2.mimecastreport.com", mid="<d2@t>"),
        dmarc_from("dmarcreport@microsoft.com", subj="[Preview] Report Domain: eaa-audit.net Submitter: enterprise.protection.outlook.com",
                   mid="<d3@t>"),
        dmarc_from("MAILER-DAEMON@esa2.mx-relay.example", mid="<d4@t>"),                       # Д1
        dmarc_from("MAILER-DAEMON@tech-group.example", mid="<d5@t>"),                                   # Д1
        promo(frm="Wine Shop <bonjour@wine-shop.example>"),                                    # Д2
    ]
    spam = [mime("indexhelper <domains@indexhelper.live>", "Submit eaa-audit.net to Search Engines", "SEO", mid="<sp@t>")]
    rc, _ = go(cfg, raws, spam=spam)
    assert rc == 0
    last = log_text(cfg).strip().splitlines()[-1]
    assert "новых 14, недоставок 2, автоответов 4, отписок 0, служебных 7, живых 1" in last, last
    assert [i["store_domain"] for i in live_items(cfg)] == ["slip-shop.example"]
    res = {r[0]: r[5] for r in rows(cfg.sent_xlsx)[1:]}
    assert "нерабочий" in res["Pharmacy Shop"] and "нерабочий" in res["DIY Shop"]
    assert [r[1] for r in rows(cfg.stop_xlsx)[1:]] == ["info@pharmacy-shop.example", "sales@diy-shop.example"]


# ---------------------------------------------------------------- 12.11, по замечанию tester: файл тревог пересобирается

SLIP_TEXT = ("Bonjour, Merci de penser à nous, c’est vraiment sympa ! Je fais suivre votre dossier auprès du service concerné "
             "qui s’occupera de vous recontacter. Je vous souhaite une belle journée. Toujours disponible, toujours attentif, "
             "même sans café. (Agent AI) On Mon, Nov 09 2026, at 07:11 PM, Slip Shop <bonjour@slip-shop.example> wrote: "
             "> Bonjour hello, > > Message bien reçu 🫡 > Nous sommes en train d'examiner votre demande avec attention.")


def entry(frm, subj, text, mid, note=None, store_domain=""):
    e = {"time": "2026-11-10 00:00:16", "from": frm, "store_domain": store_domain or frm.split("@")[-1].strip(">"),
         "subject": subj, "text": text, "message_id": mid, "webmail_hint": "https://privateemail.com/appsuite/ — x",
         "alert": True, "folder": "INBOX"}
    if note:
        e["note"] = note
    return e


def state_12_11():
    """Живой файл как его записал старый код 12.11 17:35: 4 записи, все alert=true."""
    return [
        entry("Slip Shop <bonjour@slip-shop.example>", "Accessibilité de Slip Shop : 19 éléments à corriger", SLIP_TEXT,
              "<ticket.0001@gorgias.com>"),
        entry("MAILER-DAEMON@esa2.mx-relay.example", "Report Domain: eaa-audit.net Submitter: esa2.mx-relay.example Report-ID: <rid-0001@esa2.mx-relay.example>",
              "Find attached the DMARC Aggregate Report.", "rid-0001@esa2.mx-relay.example", note="недоставка не разобрана"),
        entry("MAILER-DAEMON@tech-group.example", "Report Domain: eaa-audit.net Submitter: mail.tech-group.example Report-ID: <rid-0002@mail.tech-group.example>",
              "Find attached the DMARC Aggregate Report.", "rid-0002@mail.tech-group.example", note="недоставка не разобрана"),
        entry("Wine Shop <bonjour@wine-shop.example>", "Vous risquez d'y prendre goût 😉",
              "https://www.wine-shop.example/pages/15-ans Chez Wine Shop, on déniche des vins [M'abonner](https://x)",
              "<PROMO0001@klaviyomail.com>"),
    ]


def put_live(cfg, items):
    cfg.live_json.parent.mkdir(parents=True, exist_ok=True)
    cfg.live_json.write_text(json.dumps({"updated": "2026-11-12 17:35:57", "alert": True, "items": items}, ensure_ascii=False),
                             encoding="utf-8")


def test_rebuild_removes_false_alerts_on_run_without_new_mail(env):
    put_live(env, state_12_11())
    rc, _ = go(env, [])
    assert rc == 0
    d = json.loads(env.live_json.read_text(encoding="utf-8"))
    assert [i["store_domain"] for i in d["items"]] == ["slip-shop.example"]          # DMARC ×2 и рассылка ушли
    assert d["items"][0]["forwarded"] is True and d["items"][0]["alert"] is False
    assert d["alert"] is False
    text = log_text(env)
    assert text.count("запись убрана из eaa_mail.json") == 3 and "klaviyo" not in text.lower()


def test_rebuild_is_idempotent_and_keeps_human_alert(env):
    items = state_12_11() + [entry("Jean <jean@live-shop.example>", "Re: Accessibility check of Live Shop",
                                   "Bonjour, pouvez-vous m'envoyer un devis ?", "<j1@t>")]
    put_live(env, items)
    go(env, [])
    first = env.live_json.read_text(encoding="utf-8")
    d = json.loads(first)
    assert [i["store_domain"] for i in d["items"]] == ["slip-shop.example", "live-shop.example"]
    assert d["alert"] is True and d["items"][1]["alert"] is True and "forwarded" not in d["items"][1]
    go(env, [])
    assert env.live_json.read_text(encoding="utf-8") == first                         # второй запуск файл не меняет


def test_rebuild_keeps_unparsed_delivery_report_alert(env):
    go(env, [dsn("contact@dead-shop.example", parsed=False)])
    go(env, [])
    items = live_items(env)
    assert len(items) == 1 and items[0]["alert"] is True and items[0]["note"] == "недоставка не разобрана"


def test_rebuild_does_not_touch_missing_file(env):
    go(env, [])
    assert not env.live_json.exists()


def test_rebuild_uses_saved_headers_of_new_entries(env):
    go(env, [mime("Marie <marie@live-shop.example>", "Re: Accessibility check of Live Shop", "Merci, appelez-moi.",
                  extra="In-Reply-To: <our@x>\r\nList-Unsubscribe: <mailto:u@x.example>\r\n")])
    it = live_items(env)[0]
    assert it["meta"]["headers"]["In-Reply-To"] == "<our@x>" and "List-Unsubscribe" in it["meta"]["headers"]
    go(env, [])
    assert len(live_items(env)) == 1                      # ответ с In-Reply-To остаётся живым и после перепроверки


def test_forwarded_ai_agent_reply_is_stored_without_alert(env):
    m = mime("Le Slip <bonjour@slip-shop.example>", "Accessibilité de Slip Shop : 19 éléments à corriger", SLIP_TEXT.split(" On Mon")[0])
    go(env, [m])
    it = live_items(env)[0]
    assert it["forwarded"] is True and it["alert"] is False and json.loads(env.live_json.read_text(encoding="utf-8"))["alert"] is False


@pytest.mark.parametrize("body", [
    "Bonjour, merci pour votre rapport. Pouvez-vous m'envoyer un devis pour la correction ? Cordialement, Jean",
    "Thanks, we will get back to you next week. Can you send a quote for the fixes?",
    "Bonjour, je fais suivre à mon directeur. Appelez-moi demain au 06 00 00 00 00.",
    "Merci, ça m'intéresse. Message bien reçu, on en parle jeudi.",
], ids=["devis", "will-get-back-quote", "fais-suivre-appel", "intéresse"])
def test_human_reply_raises_alert_not_forwarded(env, body):
    go(env, [mime("Jean <jean@live-shop.example>", "Re: Accessibility check of Live Shop", body, extra="In-Reply-To: <o@x>\r\n")])
    it = live_items(env)[0]
    assert it["alert"] is True and "forwarded" not in it


WEAK_BODIES = [
    "Merci pour votre message. Je fais suivre à notre équipe.",
    "Thank you, someone will get back to you soon.",
    "Message bien reçu 🫡 Nous revenons vers vous très rapidement.",
    "Hi, I've passed this to our web team, we will get back to you.",
]
WEAK_IDS = ["je-fais-suivre", "will-get-back", "message-bien-recu", "passed-to-web-team"]


@pytest.mark.parametrize("body", WEAK_BODIES, ids=WEAK_IDS)
def test_weak_forward_phrase_without_strong_sign_is_live_with_alert(env, body):
    """ОДНО ПРАВИЛО: слабая фраза без признака машины — сомнение, значит живое с тревогой."""
    go(env, [mime("Support <help@live-shop.example>", "Re: Accessibility check of Live Shop", body)])
    it = live_items(env)[0]
    assert it["alert"] is True and "forwarded" not in it and it["note"] == "возможно передача в отдел"


@pytest.mark.parametrize("extra", ["X-GORGIAS-TICKET-ID: 1\r\n", "X-Zendesk-From-Account-Id: 7\r\n"], ids=["gorgias-hdr", "zendesk-hdr"])
@pytest.mark.parametrize("body", WEAK_BODIES, ids=WEAK_IDS)
def test_weak_forward_phrase_with_support_desk_header_is_forwarded(env, body, extra):
    go(env, [mime("Support <help@live-shop.example>", "Re: Accessibility check of Live Shop", body, extra=extra)])
    it = live_items(env)[0]
    assert it["forwarded"] is True and it["alert"] is False
    assert "передача в отдел, без тревоги [INBOX UID 1] от help@live-shop.example" in log_text(env)


def test_weak_forward_phrase_with_support_desk_message_id_is_forwarded(env):
    go(env, [mime("Support <help@live-shop.example>", "Re: Accessibility check", WEAK_BODIES[1], mid="<abc.123@shop.zendesk.com>")])
    assert live_items(env)[0]["forwarded"] is True


def test_ai_agent_signature_alone_is_forwarded(env):
    go(env, [mime("Support <help@live-shop.example>", "Re: Accessibility check", "Bonjour, ticket créé. (Agent AI)")])
    it = live_items(env)[0]
    assert it["forwarded"] is True and it["alert"] is False


@pytest.mark.parametrize("dom", ["sendgrid.net", "mandrillapp.com", "sparkpostmail.com", "mailjet.com", "brevo.com",
                                 "sendinblue.com", "hubspotemail.net"])
def test_operator_relay_message_id_is_not_bulk(env, dom):
    go(env, [mime("Anna <anna@live-shop.example>", "Votre rapport", "Bonjour, je regarde votre rapport cette semaine.",
                  mid=f"<x1.y2@{dom}>")])
    it = live_items(env)
    assert len(it) == 1 and it[0]["alert"] is True


# ---------------------------------------------------------------- урок: перепроверка не понижает живое

LONG_HUMAN = ("Bonjour, merci pour votre message, je fais suivre à notre équipe technique qui gère le site. " * 6
              + "Mais dites-moi : combien coûte l'audit complet ?")        # вопрос после 500-го знака


def live_with_alert_samples():
    return {
        "devis": mime("Jean <jean@live-shop.example>", "Re: Accessibility check of Live Shop", "Pouvez-vous m'envoyer un devis ?",
                      mid="<s1@t>"),
        "weak-phrase": mime("Anna <anna@live-shop.example>", "Re: Accessibility check", WEAK_BODIES[3], mid="<s2@t>"),
        "long-question-after-500": mime("Paul <paul@live-shop.example>", "Re: Accessibility check", LONG_HUMAN, mid="<s3@t>"),
        "store-bulk-about-us": mime("Marie <marie@slip-shop.example>", "Votre rapport", "J'ai lu votre audit d'accessibilité.",
                                    extra="List-Unsubscribe: <mailto:u@x.example>\r\n", mid="<s4@t>"),
        "dmarc-from-store": dmarc_from("Slip Shop <bonjour@slip-shop.example>", subj="Report for eaa-audit.net", text="Voir pièce jointe",
                                       attach="report.zip", mid="<s5@t>"),
        "maybe-autoreply": mime("Slip <bonjour@slip-shop.example>", "Re: Accessibility check",
                                "Votre demande a bien été transmise à notre équipe.", mid="<s6@t>"),
        "unparsed-dsn": dsn("contact@dead-shop.example", parsed=False),
        "operator-relay": mime("Anna <anna@live-shop.example>", "Votre rapport", "Je regarde.", mid="<x@sendgrid.net>"),
    }


@pytest.mark.parametrize("name", list(live_with_alert_samples()))
@pytest.mark.parametrize("drop_xlsx", [False, True], ids=["stores", "no-stores"])
def test_lesson_live_alert_survives_second_empty_run(env, name, drop_xlsx):
    go(env, [live_with_alert_samples()[name]])
    first = live_items(env)
    assert len(first) == 1 and first[0]["alert"] is True, first
    if drop_xlsx:
        env.sent_xlsx.unlink()
    go(env, [])                                                    # второй, пустой запуск: перепроверка файла
    again = live_items(env)
    assert len(again) == 1 and again[0]["alert"] is True and "forwarded" not in again[0]
    assert json.loads(env.live_json.read_text(encoding="utf-8"))["alert"] is True


def test_meta_keeps_full_text_before_quote(env):
    go(env, [mime("Paul <paul@live-shop.example>", "Re: Accessibility check", LONG_HUMAN + "\r\n\r\nLe lun. 9 nov. 2026, X a écrit :\r\n> old")])
    it = live_items(env)[0]
    assert len(it["text"]) <= 500 and it["meta"]["head"].endswith("combien coûte l'audit complet ?")
    assert it["meta"]["verdict"] == "live" and "a écrit" not in it["meta"]["head"]


def test_empty_store_list_skips_recheck_with_log(env):
    put_live(env, state_12_11())                                 # старые записи, в том числе 2 DMARC и рассылка
    env.sent_xlsx.unlink()
    go(env, [])
    assert len(live_items(env)) == 4
    assert "перепроверка eaa_mail.json пропущена: список магазинов пуст" in log_text(env)


def test_legacy_entry_is_not_dropped_without_machine_sign(env):
    put_live(env, [entry("Bob <bob@agency.example>", "Partenariat", "Bonjour, nous proposons du référencement SEO et des backlinks.",
                         "<seo@agency.example>")])
    go(env, [])
    assert len(live_items(env)) == 1 and live_items(env)[0]["alert"] is True


def test_legacy_downgrades_are_logged_with_sender_and_subject(env):
    put_live(env, state_12_11())
    go(env, [])
    text = log_text(env)
    assert "запись убрана из eaa_mail.json (отчёт DMARC): от MAILER-DAEMON@tech-group.example: «Report Domain" in text
    assert "запись убрана из eaa_mail.json (рассылка (Message-ID сервера рассылок)): от Wine Shop" in text
    assert "тревога снята, передача в отдел (подпись ИИ-агента / система поддержки): от Slip Shop" in text


# XML задачи Планировщика проверяется в test_install_task.py (шаблон + подстановка).


# ---------------------------------------------------------------- 12.11, reviewer, круг 1: Б1, Б2, В3, В4

FIX16 = Path(__file__).resolve().parent / "test_data" / "uid16_support_reply.eml"
SLIP_SUBJ = "Accessibilité de Slip Shop : 19 éléments à corriger"


def fix16_bytes() -> bytes:
    """Файл письма с CRLF независимо от того, как git переписал концы строк при checkout."""
    return FIX16.read_bytes().replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")


def real16(drop=(), add=""):
    """Настоящие заголовки INBOX UID 16 (ответ магазина через Gorgias, 09.11; домен и имя заменены)."""
    head, body = fix16_bytes().split(b"\r\n\r\n", 1)
    keep = [ln for ln in head.split(b"\r\n") if not any(ln.lower().startswith(d.lower().encode() + b":") for d in drop)]
    return b"\r\n".join(keep) + add.encode() + b"\r\n\r\n" + body


def slip_xlsx(env):
    make_xlsx(env.sent_xlsx, subjects={"Slip Shop": SLIP_SUBJ})


def test_b1_fixture_has_the_real_uid16_headers():
    h = fix16_bytes().split(b"\r\n\r\n", 1)[0].decode("utf-8")
    for name in ("In-Reply-To:", "References:", "X-GORGIAS-TICKET-ID:", "Reply-To:", "@gorgias.com>"):
        assert name in h
    assert "List-Unsubscribe" not in h and "Precedence" not in h
    # все адреса в образце — выдуманные .example, кроме нашего ящика и домена системы поддержки
    doms = {a.rsplit("@", 1)[1].rstrip(">").lower() for a in re.findall(r"[\w.+=-]+@[\w.-]+", h)}
    assert doms <= {"slip-shop.example", "em1.slip-shop.example", "eaa-audit.net", "gorgias.com", "mailcore.example"}, doms


def test_b1_real_uid16_reply_is_kept_as_forwarded(env):
    slip_xlsx(env)
    go(env, [real16()])
    it = live_items(env)
    assert len(it) == 1 and it[0]["forwarded"] is True and it[0]["store_domain"] == "slip-shop.example"


def test_b1_uid16_without_reply_headers_but_with_bulk_is_kept_by_our_subject(env):
    slip_xlsx(env)
    go(env, [real16(drop=("In-Reply-To", "References"), add="\r\nList-Unsubscribe: <mailto:u@x.example>")])
    assert len(live_items(env)) == 1 and "служебное" not in log_text(env)


@pytest.mark.parametrize("subj", ["[#4521] Re: Votre demande", "Fwd: Votre demande", "FW: Votre demande", "WG: Ihre Anfrage",
                                  "TR: Votre demande"])
def test_b1_reply_prefixes_with_bulk_headers_are_kept(env, subj):
    slip_xlsx(env)
    go(env, [real16(drop=("In-Reply-To", "References", "Subject"), add=f"\r\nSubject: {subj}\r\nPrecedence: bulk")])
    assert len(live_items(env)) == 1


def test_b1_store_bulk_human_without_re_mentioning_us_is_live_with_note(env):
    m = mime("Marie <marie@slip-shop.example>", "Votre rapport", "Bonjour, j'ai lu votre audit d'accessibilité. Appelez-moi demain.",
             extra="List-Unsubscribe: <mailto:u@x.example>\r\n")
    go(env, [m])
    it = live_items(env)
    assert len(it) == 1 and it[0]["alert"] is True and it[0]["note"] == "возможно рассылка"


def test_b1_store_bulk_without_anything_about_us_stays_service(env):
    go(env, [promo()])
    assert not env.live_json.exists() and "служебное [INBOX UID 1]" in log_text(env)


NO_SUCH_USER_TEXT = ("Delivery to the following recipient failed permanently: contact@dead-shop.example "
                "Technical details of permanent failure: 550 No Such User Here")
O365_REJECTED_DIAG = ("smtp; 550 5.4.1 Recipient address rejected: Access denied. For more information see "
                "https://aka.ms/EXOSmtpErrors [DB1PEPF000509E5.eurprd03.prod.outlook.com 2026-11-09T19:02:57.233Z 08DF3BEEB577C59E]")


@pytest.mark.parametrize("diag,text", [
    ("smtp; 550 No Such User Here", NO_SUCH_USER_TEXT),
    (O365_REJECTED_DIAG, NO_SUCH_USER_TEXT.replace("550 No Such User Here", O365_REJECTED_DIAG[6:])),
], ids=["no-such-user", "o365-recipient-rejected"])
def test_b2_real_live_cases_stay_hard_and_are_logged(env, diag, text):
    cell, stop = hard_case(env, diag=diag, text=text)
    assert "нерабочий" in cell and "недоставка временно/отказ" not in cell and len(stop) == 2
    assert "недоставка жёсткая [INBOX UID 1]" in log_text(env) and "contact@dead-shop.example -> Не_писать" in log_text(env)


def test_b2_block_word_at_150th_char_is_soft(env):
    diag = "smtp; 550 No such user here. "
    diag += "x" * (150 - len(diag)) + " blocked by local policy"
    assert diag.index("blocked") >= 150
    cell, stop = hard_case(env, diag=diag, text="Delivery failed")
    assert "недоставка временно/отказ" in cell and "нерабочий" not in cell and stop == []


@pytest.mark.parametrize("diag", [
    "smtp; 550 Sender verify failed",
    "smtp; 550 Your IP 1.2.3.4 is listed in black list",
    "smtp; 550 Message rejected by Anti-Spoofing policy",
    "smtp; 550 Access denied",
    "smtp; 550 Mailbox unavailable",
    "smtp; 553 Requested action not taken",
], ids=["sender-verify", "black-list", "anti-spoofing", "access-denied", "mailbox-unavailable", "553-no-phrase"])
def test_b2_other_550_is_soft(env, diag):
    cell, stop = hard_case(env, diag=diag, text="Delivery failed")
    assert "недоставка временно/отказ" in cell and "нерабочий" not in cell and stop == []


def test_b2_sender_status_518_is_not_dead_address(env):
    cell, stop = hard_case(env, status="5.1.8", diag="smtp; 550 5.1.8 Sender address rejected", text="Delivery failed")
    assert "нерабочий" not in cell and stop == []


@pytest.mark.parametrize("diag", ["smtp; 550 user unknown", "smtp; 550 Adresse introuvable", "smtp; 551 l'adresse n'existe pas",
                                  "smtp; 550 mailbox not found", "smtp; 553 invalid recipient"])
def test_b2_white_list_phrases_are_hard(env, diag):
    cell, stop = hard_case(env, diag=diag, text="Delivery failed")
    assert "нерабочий" in cell and len(stop) == 2


def test_b2_dead_mark_replaces_soft_mark_in_cell(env):
    raws = [dsn("contact@dead-shop.example", status="5.0.0", diag="smtp; 550 Access denied", text="Delivery failed"),
            dsn("contact@dead-shop.example", status="5.0.0", diag="smtp; 550 No Such User Here")]
    go(env, raws[:1])                                                           # сначала мягкая метка
    assert "недоставка временно/отказ" in rows(env.sent_xlsx)[1][5]
    go(env, raws)                                                               # потом жёсткая (UID 2)
    cell = rows(env.sent_xlsx)[1][5]
    assert "нерабочий" in cell and "недоставка временно/отказ" not in cell


def test_v3_store_newsletter_with_unsubscribe_first_line_is_service_not_stop(env):
    m = mime("Slip Shop <news@slip-shop.example>", "Nos nouveautés de novembre",
             "Se désinscrire | Voir en ligne\r\nDécouvrez la nouvelle collection.", extra="List-Unsubscribe: <mailto:u@x.example>\r\n")
    go(env, [m])
    assert not env.stop_xlsx.exists() and not env.live_json.exists()
    assert "служебное [INBOX UID 1] от news@slip-shop.example" in log_text(env)


def test_v4_client_asks_for_rapport_agrege_is_live(env):
    slip_xlsx(env)
    m = mime("Marie <marie@slip-shop.example>", "Re: " + SLIP_SUBJ,
             "Bonjour, pouvez-vous nous envoyer le rapport agrégé de l'audit ?", extra="In-Reply-To: <o@x>\r\n")
    go(env, [m])
    it = live_items(env)
    assert len(it) == 1 and it[0]["alert"] is True and it[0]["note"] == "возможно отчёт DMARC"


def test_v4_store_report_zip_is_live_not_silent(env):
    go(env, [dmarc_from("Slip Shop <bonjour@slip-shop.example>", subj="Report for eaa-audit.net", text="Voir pièce jointe",
                        attach="report.zip")])
    it = live_items(env)
    assert len(it) == 1 and it[0]["note"] == "возможно отчёт DMARC"


# ---------------------------------------------------------------- 13.11, reviewer: признак машины только однозначный

AGENT_QUOTE = ("Bonjour, merci de penser à nous ! Je fais suivre votre dossier auprès du service concerné. "
               "Toujours disponible, même sans café. (Agent AI)")


def html_mail(frm, subj, html, extra="", mid=None):
    return mime(frm, subj, html, extra=extra, mid=mid, ctype="text/html; charset=utf-8")


@pytest.mark.parametrize("own", ["Merci, appelez-moi demain matin.", "Merci, je regarde ça cette semaine avec mon équipe."],
                         ids=["with-human-word", "neutral"])
def test_r4_html_human_reply_quoting_agent_ai_raises_alert(env, own):
    html = (f"<div dir='ltr'><p>{own}</p><p>Marie</p></div><div class='gmail_quote'><div>Le lun. 9 nov. 2026, Slip Shop "
            f"&lt;bonjour@slip-shop.example&gt; a écrit :</div><blockquote>{AGENT_QUOTE}</blockquote></div>")
    go(env, [html_mail("Marie <marie@slip-shop.example>", "Re: Accessibility check", html,
                       extra="X-GORGIAS-TICKET-ID: 5\r\nIn-Reply-To: <o@x>\r\n")])
    it = live_items(env)[0]
    assert it["alert"] is True and "forwarded" not in it


def test_r4_agent_ai_not_at_the_end_is_not_a_machine_sign(env):
    body = "(Agent AI) m'a répondu hier. " + "Je reprends le sujet moi-même avec notre développeur web, " * 5 + "à bientôt."
    assert len(body) - body.index("(Agent AI)") > 200
    go(env, [mime("Paul <paul@live-shop.example>", "Re: Accessibility check", body)])
    assert live_items(env)[0]["alert"] is True


def test_r4_weak_phrase_gorgias_and_question_at_900th_char_raises_alert(env):
    body = "Bonjour, je fais suivre votre dossier au service concerné. " + "Nous regardons le rapport en détail. " * 23
    body += "Pouvez-vous nous dire combien coûte l'audit complet ?"
    assert body.index("?") >= 900
    go(env, [mime("Support <help@live-shop.example>", "Re: Accessibility check", body, extra="X-GORGIAS-TICKET-ID: 7\r\n")])
    it = live_items(env)[0]
    assert it["alert"] is True and "forwarded" not in it


def test_r4_ai_agent_signature_at_the_end_still_forwarded(env):
    go(env, [mime("Support <help@live-shop.example>", "Re: Accessibility check", AGENT_QUOTE)])
    assert live_items(env)[0]["forwarded"] is True


@pytest.mark.parametrize("extra,subj", [
    ("List-Unsubscribe: <mailto:u@x.example>\r\nIn-Reply-To: <o@x>\r\n", "Votre rapport"),
    ("List-Unsubscribe: <mailto:u@x.example>\r\nReferences: <o@x>\r\n", "Votre rapport"),
    ("Precedence: bulk\r\n", "Re: Accessibility check"),
    ("", "RE: backlinks et SEO pour votre site"),
], ids=["in-reply-to", "references", "re-bulk", "re-seo-words"])
def test_r4_empty_store_list_reply_is_live_not_service(env, extra, subj):
    env.sent_xlsx.unlink()
    go(env, [mime("Anna <anna@unknown-shop.example>", subj, "Bonjour, je regarde votre rapport.", extra=extra)])
    it = live_items(env)
    assert len(it) == 1 and it[0]["alert"] is True


def test_r4_empty_store_list_reply_in_spam_folder_is_live(env):
    env.sent_xlsx.unlink()
    go(env, [], spam=[mime("Anna <anna@unknown-shop.example>", "Votre rapport", "Bonjour.", extra="References: <o@x>\r\n")])
    assert len(live_items(env)) == 1


def test_r4_empty_store_list_newsletter_without_reply_sign_stays_service(env):
    env.sent_xlsx.unlink()
    go(env, [promo(frm="News <news@unknown-shop.example>")])
    assert not env.live_json.exists()
