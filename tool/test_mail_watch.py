"""Тесты mail_watch.py на подменённом ящике (сети нет, живого пароля нет). Запуск: python -m pytest tool/test_mail_watch.py"""
import json
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


def dsn(addr):
    b = "BOUNDARY1"
    return (f"From: MAILER-DAEMON@mail.privateemail.com\r\nTo: hello@eaa-audit.net\r\n"
            f"Subject: Undelivered Mail Returned to Sender\r\nDate: Tue, 10 Nov 2026 09:00:00 +0000\r\n"
            f"Message-ID: <dsn1@test>\r\nMIME-Version: 1.0\r\n"
            f"Content-Type: multipart/report; report-type=delivery-status; boundary={b}\r\n\r\n"
            f"--{b}\r\nContent-Type: text/plain\r\n\r\nThis is the mail system. <{addr}>: host said: 550 No Such User Here\r\n"
            f"--{b}\r\nContent-Type: message/delivery-status\r\n\r\nReporting-MTA: dns; mail.privateemail.com\r\n\r\n"
            f"Final-Recipient: rfc822; {addr}\r\nAction: failed\r\nStatus: 5.1.1\r\n"
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

    def __init__(self, raws, validity=111):
        self.raws, self.validity = raws, validity
        self.readonly = None
        self.cmds = []
        self.logged_out = False

    def __getattr__(self, name):
        if name in self.forbidden:
            raise AssertionError(f"запрещённая команда IMAP: {name}")
        raise AttributeError(name)

    def login(self, u, p):
        return "OK", []

    def select(self, box="INBOX", readonly=False):
        self.readonly = readonly
        return "OK", [b"%d" % len(self.raws)]

    def response(self, code):
        return code, [str(self.validity).encode()]

    def uid(self, cmd, *args):
        self.cmds.append((cmd, args))
        n = len(self.raws)
        if cmd == "SEARCH":
            if args[1] == "UID":
                lo = int(args[2].split(":")[0])
                ids = [u for u in range(1, n + 1) if u >= lo] or [n]
            else:
                ids = list(range(1, n + 1))
            return "OK", [b" ".join(b"%d" % i for i in ids)]
        if cmd == "FETCH":
            assert "BODY.PEEK" in args[1], "чтение без PEEK ставит флаг прочитано"
            return "OK", [(b"1 (BODY[] {1}", self.raws[int(args[0]) - 1]), b")"]
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


def go(cfg, raws, validity=111):
    box = FakeIMAP(raws, validity)
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
    assert "автоответ" in res["Slip Shop"]                # «передано в отдел» = автоответ
    assert res["Dmarc Shop"] == "отправлено" and res["Quiet Shop"] == "отправлено"
    assert res["Live Shop"] == "отправлено"               # живой xlsx не трогает
    assert cfg.sent_xlsx.with_name(cfg.sent_xlsx.name + ".bak").exists()
    stop = rows(cfg.stop_xlsx)
    assert stop[0] == ["Магазин", "Адрес", "Причина", "Дата"]
    assert [r[1] for r in stop[1:]] == ["info@unsub-shop.example"] and stop[1][0] == "Unsub Shop"
    live = json.loads(cfg.live_json.read_text(encoding="utf-8"))
    assert live["alert"] is True and len(live["items"]) == 1
    it = live["items"][0]
    assert it["alert"] is True and "jean.dupont@live-shop.example" in it["from"]
    assert it["store_domain"] == "live-shop.example" and it["subject"].startswith("Re: Accessibility")
    assert "devis" in it["text"] and len(it["text"]) <= 500 and it["message_id"] and "privateemail.com" in it["webmail_hint"]
    last = cfg.log_file.read_text(encoding="utf-8").strip().splitlines()[-1]
    assert "новых 6, недоставок 1, автоответов 2, отписок 1, служебных 1, живых 1" in last


def test_rerun_does_not_duplicate(env):
    cfg = env
    raws = six()
    go(cfg, raws)
    x1, l1, s1 = rows(cfg.sent_xlsx), cfg.live_json.read_text(encoding="utf-8"), rows(cfg.stop_xlsx)
    rc, box = go(cfg, raws)
    assert rc == 0
    assert rows(cfg.sent_xlsx) == x1 and rows(cfg.stop_xlsx) == s1
    assert cfg.live_json.read_text(encoding="utf-8").count('"message_id"') == 1
    assert "новых 0," in cfg.log_file.read_text(encoding="utf-8").strip().splitlines()[-1]
    assert box.cmds[0][1][:2] == (None, "UID") and box.cmds[0][1][2] == "7:*"   # по UID с прошлого запуска
    # состояние потеряно: те же письма не дают второго «живого» (дедуп по Message-ID) и не дублируют пометки
    cfg.state_json.unlink()
    go(cfg, raws)
    assert cfg.live_json.read_text(encoding="utf-8").count('"message_id"') == 1
    assert rows(cfg.sent_xlsx) == x1 and len(rows(cfg.stop_xlsx)) == len(s1)


def test_new_mail_after_first_run(env):
    cfg = env
    raws = six()
    go(cfg, raws)
    extra = mime("Boss <boss@quiet-shop.example>", "Re: check", "Intéressant, appelez-moi.")
    go(cfg, raws + [extra])
    items = json.loads(cfg.live_json.read_text(encoding="utf-8"))["items"]
    assert len(items) == 2 and items[1]["store_domain"] == "quiet-shop.example"


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
    assert rc == 0 and len(json.loads(env.live_json.read_text(encoding="utf-8"))["items"]) == 1


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
    assert set(diff) == {pick["dead"], pick["auto"], pick["slip"]}
    assert "нерабочий" in diff[pick["dead"]] and "автоответ" in diff[pick["auto"]] and "автоответ" in diff[pick["slip"]]
    assert [r[1] for r in rows(cfg.stop_xlsx)[1:]] == [pick["unsub"]]
    items = json.loads(cfg.live_json.read_text(encoding="utf-8"))["items"]
    assert len(items) == 1 and items[0]["store_domain"] == "ldlc.com"
