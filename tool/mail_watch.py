"""Слежение за ответами на письма рассылки (ящик hello@eaa-audit.net). Короткий запуск и выход.

    python mail_watch.py            — один проход: новые письма -> разбор -> пометки/тревога -> выход
    python mail_watch.py save       — владелец один раз: пароль из буфера обмена -> DPAPI -> буфер очищается
                                      (запускать через «Сохранить_пароль_почты.bat»)
    python mail_watch.py check      — войти в ящик (без чтения писем) и сказать «пароль верный / не подошёл»

Ящик открывается ТОЛЬКО на чтение (select readonly + BODY.PEEK): письма не удаляются, не перемещаются,
флаг «прочитано» не ставится. Пароль нигде не печатается и не пишется (в сообщениях об ошибках вырезается).

Коды выхода: 0 — ок; 1 — ошибка сети/IMAP/чтения письма; 2 — нет пароля; 3 — не удалось записать xlsx/json/состояние;
4 — уже идёт другой запуск.
Главное правило: курсор (last_uid) двигается только до подтверждённого результата; живой ответ не теряется.
Пути настраиваются переменными EAA_MAIL_* (см. default_config), нужны тестам и копиям.
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import email
import imaplib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from email import policy
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HOST, PORT, USER = "mail.privateemail.com", 993, "hello@eaa-audit.net"
FIRST_RUN_SINCE = "09-Nov-2026"
WEBMAIL = "https://privateemail.com/appsuite/"
STORE = Path(os.environ.get("LOCALAPPDATA", ".")) / "OfficeKeys" / "eaa_mail.dpapi"
MARK_DEAD, MARK_AUTO = "нерабочий", "автоответ"

C1 = Path("C:/!Projects/Office/workspace/eaa/круг1")


@dataclass
class Config:
    sent_xlsx: Path = C1 / "Отправлено.xlsx"
    stop_xlsx: Path = C1 / "Не_писать.xlsx"
    live_json: Path = Path("C:/!Projects/Office/data/live/eaa_mail.json")
    state_json: Path = ROOT / "state" / "mail_watch.json"
    log_file: Path = ROOT / "logs" / "mail_watch.log"
    timeout: int = 30
    host: str = HOST
    extra: dict = field(default_factory=dict)


def default_config() -> Config:
    c = Config()
    for attr, env in (("sent_xlsx", "EAA_MAIL_SENT_XLSX"), ("stop_xlsx", "EAA_MAIL_STOP_XLSX"),
                      ("live_json", "EAA_MAIL_LIVE_JSON"), ("state_json", "EAA_MAIL_STATE"),
                      ("log_file", "EAA_MAIL_LOG")):
        if os.environ.get(env):
            setattr(c, attr, Path(os.environ[env]))
    return c


# ---------------------------------------------------------------- пароль (DPAPI)

class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(data: bytes, protect: bool) -> bytes:
    src = _Blob(len(data), ctypes.cast(ctypes.create_string_buffer(data, len(data)), ctypes.POINTER(ctypes.c_char)))
    out = _Blob()
    fn = ctypes.windll.crypt32.CryptProtectData if protect else ctypes.windll.crypt32.CryptUnprotectData
    if not fn(ctypes.byref(src), None, None, None, None, 0x1, ctypes.byref(out)):
        raise OSError("Windows не смог зашифровать или расшифровать пароль")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out.pbData)


def get_password():
    try:
        return _dpapi(STORE.read_bytes(), False).decode().strip() if STORE.exists() else None
    except OSError:
        return None


def save_password() -> int:
    r = subprocess.run(["powershell", "-NoProfile", "-Command", "Get-Clipboard -Raw"],
                       capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30)
    pw = (r.stdout or "").strip()
    if not pw or "\n" in pw or "\r" in pw:
        print("В буфере обмена нет пароля (одна строка). Скопируй пароль ящика hello@eaa-audit.net и запусти снова.")
        return 1
    try:
        STORE.parent.mkdir(parents=True, exist_ok=True)
        STORE.write_bytes(_dpapi(pw.encode(), True))
    finally:
        subprocess.run(["powershell", "-NoProfile", "-Command", "Set-Clipboard -Value ' '"], timeout=30)
    print("Пароль сохранён зашифрованным, буфер обмена очищен.")
    return 0


AUTH_FAIL = re.compile(r"authenticat|login failed|invalid credentials|bad credentials|wrong password", re.I)


def check_password(cfg=None, imap_factory=None, password=None) -> int:
    """Только вход в ящик (письма не открываются). Пароль не печатается."""
    cfg = cfg or default_config()
    pw = password if password is not None else get_password()
    if not pw:
        print("Пароль не сохранён: запусти Сохранить_пароль_почты.bat")
        return 2
    try:
        M = (imap_factory or connect)(cfg, pw)
    except imaplib.IMAP4.abort:
        print("Нет связи с почтой, попробуйте позже. Пароль не проверен.")
        return 3
    except imaplib.IMAP4.error as e:
        if AUTH_FAIL.search(str(e)):
            print("Пароль НЕ подошёл: сервер отказал во входе. Скопируй пароль ещё раз и запусти снова.")
            return 1
        print("Нет связи с почтой, попробуйте позже. Пароль не проверен.")
        return 3
    except Exception:  # noqa: BLE001  (сеть, таймаут, DNS)
        print("Нет связи с почтой, попробуйте позже. Пароль не проверен.")
        return 3
    try:
        M.logout()
    except Exception:  # noqa: BLE001
        pass
    print("Пароль верный: вход в ящик выполнен.")
    return 0


# ---------------------------------------------------------------- журнал

def log(cfg: Config, line: str, secret: str | None = None) -> None:
    if secret:
        line = line.replace(secret, "***")
    cfg.log_file.parent.mkdir(parents=True, exist_ok=True)
    with cfg.log_file.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {line}\n")


# ---------------------------------------------------------------- классификация (все правила здесь)

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

DSN_FROM = re.compile(r"mailer-daemon|postmaster|mail delivery subsystem", re.I)
DSN_SUBJ = re.compile(
    r"undeliver|delivery status notification|delivery failure|mail delivery failed|returned mail|"
    r"failure notice|delivery has failed|could not be delivered|not delivered|"
    r"adresse introuvable|[ée]chec de (la )?(livraison|distribution)|non remis|message non distribu|"
    r"impossible de (vous )?(livrer|remettre)|n'a pas pu [êe]tre (remis|distribu|livr)", re.I)
DMARC_SUBJ = re.compile(r"report domain:|dmarc|rapport agr[ée]g|aggregate report", re.I)
AUTO_SUBJ = re.compile(
    r"auto(matic)?[- ]?(reply|response)|autoreply|out of office|r[ée]ponse automatique|"
    r"accus[ée] de r[ée]ception|absence du bureau|absent du bureau|ne pas r[ée]pondre", re.I)
AUTO_TEXT = re.compile(
    r"we(?:'ve| have)? received your (request|message|email|e-mail)|your (request|message|email) (has been|was) "
    r"(received|forwarded|passed)|has been forwarded to|thank you for contacting|"
    r"accus[ée] de r[ée]ception|nous avons bien re[çc]u|votre (demande|message|mail|e-mail) (a|vient d')\s*"
    r"(bien )?[ée]t[ée]? (re[çc]u|enregistr|transmis|pris en compte)|a bien [ée]t[ée] (transmis|re[çc]u|enregistr)|"
    r"transmis(e)? (à|au) (notre|l[’']|le )?(service|[ée]quipe|d[ée]partement)|"
    r"(ceci|ce message) est un (message|mail|e-mail|courriel) automatique|this is an automated", re.I)
UNSUB_TEXT = re.compile(
    r"unsubscribe|rem.ve me|do not (contact|email|write)|don't (contact|email|write)|stop (emailing|writing|contacting)|"
    r"take me off|ne plus (recevoir|me contacter|nous contacter|nous [ée]crire|m[’']?[ée]crire)|"
    r"ne m[’']?[ée]crivez plus|ne nous [ée]crivez plus|d[ée]sabonn|d[ée]sinscri|retirer (mon|notre) (adresse|mail)|"
    r"supprimer (mon|notre) (adresse|mail)", re.I)
SEO_TEXT = re.compile(
    r"\bSEO\b|backlink|rank (your|higher)|first page of google|premi[èe]re page (de|sur) google|"
    r"r[ée]f[ée]rencement|guest post|link building|boost your (traffic|sales|ranking)", re.I)
SERVICE_DOMAINS = ("namecheap.com", "privateemail.com", "github.com", "google.com", "googlemail.com",
                   "eaa-audit.net", "dmarcian.com", "postmarkapp.com")


GENERIC_DOMAINS = frozenset((
    "gmail.com", "googlemail.com", "outlook.com", "outlook.fr", "outlook.ie", "hotmail.com", "hotmail.fr",
    "hotmail.co.uk", "live.com", "live.fr", "msn.com", "yahoo.com", "yahoo.fr", "yahoo.ie", "yahoo.co.uk", "ymail.com",
    "orange.fr", "wanadoo.fr", "free.fr", "sfr.fr", "neuf.fr", "bbox.fr", "laposte.net", "icloud.com", "me.com",
    "aol.com", "proton.me", "protonmail.com", "gmx.com", "gmx.de", "gmx.net", "web.de", "t-online.de", "eircom.net",
    "iol.ie", "mail.com", "zoho.com", "yandex.com", "mail.ru"))


def _domain(addr: str) -> str:
    return addr.rsplit("@", 1)[-1].lower().strip(">")


def _same_domain(a: str, b: str) -> bool:
    return a == b or a.endswith("." + b) or b.endswith("." + a)


def store_match(sender: str, stores: dict):
    """Строка списка отправленных для отправителя: полный адрес; по домену — кроме общих почтовых доменов."""
    if not sender:
        return None
    if sender in stores:
        return sender
    sd = _domain(sender)
    if sd in GENERIC_DOMAINS:
        return None
    for a in stores:
        ad = _domain(a)
        if ad not in GENERIC_DOMAINS and _same_domain(sd, ad):
            return a
    return None


QUOTE_RE = re.compile(
    r"(?im)^[ \t]*(?:>|-{2,}\s*(?:original message|message d.origine|forwarded)|_{8,}|"
    r"on\b[^\n]{0,200}?(?:\n[^\n]{0,120})?wrote:|le\b[^\n]{0,200}?(?:\n[^\n]{0,120})?a [ée]crit\s*:)")


def strip_quote(body: str) -> str:
    """Текст письма до цитаты («On … wrote:», «Le … a écrit», «-----Original Message»)."""
    m = QUOTE_RE.search(body)
    return body[:m.start()] if m else body


def html_to_text(s: str) -> str:
    s = re.sub(r"(?is)<(script|style).*?</\1>", " ", s)
    s = re.sub(r"(?s)<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _text_of(part) -> str:
    try:
        t = part.get_content()
    except Exception:  # noqa: BLE001
        return ""
    return html_to_text(t) if part.get_content_type() == "text/html" else str(t)


def message_texts(msg) -> tuple[str, str]:
    """(текст самого письма, всё остальное: заголовки и текст вложенных писем/вложений) — для адресов и фраз."""
    bp = msg.get_body(preferencelist=("plain", "html"))
    body = _text_of(bp) if bp is not None else ""
    rest = []
    for part in msg.walk():
        rest.extend(f"{k}: {v}" for k, v in part.items())
        if not part.is_multipart() and part is not bp and part.get_content_maintype() == "text":
            rest.append(_text_of(part))
    return body, chr(10).join(rest)


def has_attachment_ext(msg, exts) -> bool:
    for part in msg.walk():
        name = (part.get_filename() or "").lower()
        if name.endswith(exts):
            return True
    return False


@dataclass
class Verdict:
    kind: str                 # dsn | auto | unsub | service | live
    reason: str = ""


def classify(msg, body: str, rest: str, stores: dict) -> Verdict:
    """stores: {адрес_магазина(lower): строка}. Сомнительное — live."""
    subj = str(msg.get("Subject", "") or "")
    frm = getaddresses([str(msg.get("From", "") or "")])
    sender = frm[0][1].lower() if frm else ""
    sdom = _domain(sender) if sender else ""
    ctype = msg.get_content_type()
    report_type = str(msg.get_param("report-type") or "").lower()
    head = strip_quote(body)[:800]

    # а) недоставка
    if (DSN_FROM.search(sender) or DSN_FROM.search(str(msg.get("From", "")))
            or (ctype == "multipart/report" and report_type == "delivery-status")
            or DSN_SUBJ.search(subj)):
        return Verdict("dsn", "DSN")
    # г) DMARC-отчёты
    if DMARC_SUBJ.search(subj) or ("dmarc" in sender or "dmarc" in sdom) or \
            (has_attachment_ext(msg, (".zip", ".gz", ".xml")) and re.search(r"dmarc|report", subj + sender, re.I)):
        return Verdict("service", "DMARC")
    # б) автоответ по заголовкам
    auto_sub = str(msg.get("Auto-Submitted", "") or "").strip().lower()
    if ((auto_sub and auto_sub != "no") or msg.get("X-Autoreply") or msg.get("X-Autorespond")
            or str(msg.get("Precedence", "")).strip().lower() == "auto_reply"
            or AUTO_SUBJ.search(subj)):
        return Verdict("auto", "заголовок/тема")
    if AUTO_TEXT.search(head):
        return Verdict("auto", "текст")        # без заголовка: process() запишет как живое «возможно автоответ»
    # в) отписка: только от адресата рассылки и только в начале письма (хвост «unsubscribe» бывает у любой рассылки)
    is_store = store_match(sender, stores) is not None
    if is_store and (UNSUB_TEXT.search(subj) or UNSUB_TEXT.search(head[:600])):
        return Verdict("unsub", "отписка")
    # г) служебное, рассылки, SEO
    bulk = msg.get("List-Unsubscribe") or str(msg.get("Precedence", "")).strip().lower() in ("bulk", "list", "junk")
    if any(_same_domain(sdom, d) for d in SERVICE_DOMAINS) or (bulk and not is_store) \
            or (not is_store and not msg.get("In-Reply-To") and SEO_TEXT.search(subj + " " + head)):
        return Verdict("service", "служебное/рассылка/SEO")
    return Verdict("live", "")


def find_store_row(candidates, stores: dict, sender: str = ""):
    """Адрес из списка отправленных: отправитель (адрес/домен, но не общий почтовый), иначе любой адрес из текста."""
    m = store_match(sender, stores)
    if m:
        return m
    for c in candidates:
        if c in stores:
            return c
    return None


# ---------------------------------------------------------------- xlsx

def load_stores(cfg: Config) -> dict:
    import openpyxl
    if not cfg.sent_xlsx.exists():
        return {}
    wb = openpyxl.load_workbook(cfg.sent_xlsx, read_only=True)
    ws = wb.active
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and len(row) > 1 and row[1] and "@" in str(row[1]):
            out[str(row[1]).strip().lower()] = str(row[0] or "")
    wb.close()
    return out


def _atomic_save(wb, path: Path) -> None:
    tmp = path.with_name(path.name + ".tmp")
    wb.save(tmp)
    os.replace(tmp, path)


def apply_marks(cfg: Config, marks: dict, stops: list) -> None:
    """marks {адрес: {метки}} -> столбец «Результат» (F) Отправлено.xlsx; stops [(магазин, адрес, причина)] -> Не_писать.xlsx.
    Другие столбцы не трогаем; перед записью копия .bak."""
    import openpyxl
    if marks and cfg.sent_xlsx.exists():
        wb = openpyxl.load_workbook(cfg.sent_xlsx)
        ws = wb.active
        hdr = [str(c.value or "") for c in ws[1]]
        col = (hdr.index("Результат") + 1) if "Результат" in hdr else 6
        changed = False
        for row in ws.iter_rows(min_row=2):
            addr = str(row[1].value or "").strip().lower()
            for mark in sorted(marks.get(addr, ())):
                cell = row[col - 1]
                cur = str(cell.value or "")
                if mark not in cur:
                    cell.value = f"{cur}; {mark}" if cur else mark
                    changed = True
        if changed:
            shutil.copy2(cfg.sent_xlsx, cfg.sent_xlsx.with_name(cfg.sent_xlsx.name + ".bak"))
            _atomic_save(wb, cfg.sent_xlsx)
    if stops:
        if cfg.stop_xlsx.exists():
            wb = openpyxl.load_workbook(cfg.stop_xlsx)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Не писать"
            ws.append(["Магазин", "Адрес", "Причина", "Дата"])
        have = {str(r[1] or "").strip().lower() for r in ws.iter_rows(min_row=2, values_only=True) if len(r) > 1}
        added = False
        for store, addr, why in stops:
            if addr.lower() not in have:
                ws.append([store, addr, why, datetime.now().strftime("%Y-%m-%d")])
                have.add(addr.lower())
                added = True
        if added:
            if cfg.stop_xlsx.exists():
                shutil.copy2(cfg.stop_xlsx, cfg.stop_xlsx.with_name(cfg.stop_xlsx.name + ".bak"))
            cfg.stop_xlsx.parent.mkdir(parents=True, exist_ok=True)
            _atomic_save(wb, cfg.stop_xlsx)


# ---------------------------------------------------------------- состояние и живые

def _read_json(p: Path, default):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return default


def _write_json(p: Path, data) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_name(p.name + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, p)


def live_entry(msg, body: str, store_domain: str, note: str = "") -> dict:
    frm = str(msg.get("From", "") or "")
    subj = str(msg.get("Subject", "") or "")
    try:
        when = parsedate_to_datetime(str(msg.get("Date"))).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:  # noqa: BLE001
        when = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    e = {
        "time": when, "from": frm, "store_domain": store_domain, "subject": subj,
        "text": re.sub(r"\s+", " ", body).strip()[:500],
        "message_id": str(msg.get("Message-ID", "") or "").strip(),
        "webmail_hint": f"{WEBMAIL} — Входящие, искать «{subj}» от {frm}",
        "alert": True,
    }
    if note:
        e["note"] = note
        e["webmail_hint"] = f"[{note}] " + e["webmail_hint"]
    return e


def _live_ids(cfg: Config) -> set:
    d = _read_json(cfg.live_json, {})
    return {i.get("message_id") for i in d.get("items", []) if isinstance(i, dict)} if isinstance(d, dict) else set()


def write_live(cfg: Config, lives: list) -> bool:
    """Дописать живые в eaa_mail.json. Нечитаемый файл не перезаписываем: копия .corrupt, журнал, False."""
    if not lives:
        return True
    items = []
    if cfg.live_json.exists():
        try:
            prev = json.loads(cfg.live_json.read_text(encoding="utf-8"))
            if not isinstance(prev, dict) or not isinstance(prev.get("items", []), list):
                raise ValueError("не та схема")
            items = prev.get("items", [])
        except (ValueError, UnicodeDecodeError, OSError) as e:
            try:
                dst = cfg.live_json.with_name(cfg.live_json.name + ".corrupt")
                if not dst.exists():                  # копия одна: каждый час новую не плодим
                    shutil.copy2(cfg.live_json, dst)
            except OSError:
                pass
            log(cfg, f"eaa_mail.json нечитаем ({type(e).__name__}), не перезаписываю; копия .corrupt рядом")
            return False
    have = {i.get("message_id") for i in items if isinstance(i, dict)}
    items += [e for e in lives if not e["message_id"] or e["message_id"] not in have]
    try:
        _write_json(cfg.live_json, {"updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "alert": any(i.get("alert") for i in items if isinstance(i, dict)), "items": items})
    except OSError as e:
        log(cfg, f"не записан eaa_mail.json: {e}")
        return False
    return True


def _touch_lock(cfg: Config) -> None:
    """Замок считается «старым» от последней активности: обновляем его время на каждом письме."""
    p = cfg.state_json.with_name("mail_watch.lock")
    try:
        now = time.time()
        os.utime(p, (now, now))
    except OSError:
        pass


def _acquire_lock(cfg: Config):
    p = cfg.state_json.with_name("mail_watch.lock")
    p.parent.mkdir(parents=True, exist_ok=True)
    for _ in range(2):
        try:
            fd = os.open(p, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            return p
        except FileExistsError:
            try:
                if time.time() - p.stat().st_mtime > 600:      # замок от упавшего запуска
                    p.unlink()
                    continue
            except OSError:
                pass
            return None
    return None


# ---------------------------------------------------------------- IMAP

def connect(cfg: Config, password: str):
    M = imaplib.IMAP4_SSL(cfg.host, PORT, timeout=cfg.timeout)
    M.login(USER, password)
    return M


SPAM_NAME = re.compile(r"^(inbox[./])?(spam|junk)([ ._-]?(e-?mail|mail|folder))?$", re.I)


def find_spam_folder(M):
    """Имя папки спама через LIST: по флагу Junk, иначе по имени Spam/Junk/INBOX.Spam. Нет — None."""
    try:
        typ, data = M.list()
    except Exception:  # noqa: BLE001
        return None
    if typ != "OK":
        return None
    by_name = None
    for line in data or []:
        if not isinstance(line, (bytes, str)):
            continue
        t = line.decode("utf-8", "replace") if isinstance(line, bytes) else line
        m = re.match(r'\((?P<flags>[^)]*)\)\s+(?:"[^"]*"|NIL)\s+(?P<name>.+)$', t.strip())
        if not m:
            continue
        name = m.group("name").strip().strip('"')
        if chr(92) + "junk" in m.group("flags").lower():
            return name
        if SPAM_NAME.match(name) and by_name is None:
            by_name = name
    return by_name


def _quote_folder(name: str) -> str:
    """imaplib не квотирует имена папок сам: с пробелом/скобками/кавычкой нужны двойные кавычки."""
    if name.startswith('"') or not re.search(r'[\s()"{%*\\]', name):
        return name
    return '"' + name.replace("\\", "\\\\").replace('"', '\\"') + '"'


@dataclass
class Fetched:
    got: list                   # [(uid, bytes)] — непрерывный префикс списка UID, без пропусков
    validity: int | None
    reset: bool = False         # UIDVALIDITY изменился: курсор этой папки обнуляется
    failed: str = ""            # причина, если чтение оборвалось (курсор двигается только до got)


def fetch_new(M, folder: str, last_uid: int, known_validity=None, tick=None) -> Fetched:
    """Только чтение: select(readonly) и BODY.PEEK. Ошибка/пустой ответ FETCH — стоп: письма после него не берём."""
    typ, _ = M.select(_quote_folder(folder), readonly=True)
    if typ != "OK":
        raise imaplib.IMAP4.error(f"не открылась папка {folder}")
    validity = None
    try:
        v = M.response("UIDVALIDITY")[1]
        validity = int(v[0]) if v and v[0] else None
    except Exception:  # noqa: BLE001
        pass
    reset = bool(validity and known_validity and validity != known_validity)
    if reset:
        last_uid = 0  # ящик пересоздан: читаем заново с 09.11, живые не дублируются по Message-ID
    if last_uid:
        typ, data = M.uid("SEARCH", None, "UID", f"{last_uid + 1}:*")
    else:
        typ, data = M.uid("SEARCH", None, "SINCE", FIRST_RUN_SINCE)
    if typ != "OK":
        raise imaplib.IMAP4.error(f"SEARCH в папке {folder} не удался")
    uids = sorted(int(x) for x in (data[0] or b"").split())
    uids = [u for u in uids if u > last_uid]
    out, failed = [], ""
    for u in uids:
        try:
            typ, d = M.uid("FETCH", str(u), "(BODY.PEEK[])")
            raw = next((x[1] for x in (d or []) if isinstance(x, tuple)), None)
        except Exception as e:  # noqa: BLE001
            typ, raw = "ERR", None
            failed = f"UID {u}: {type(e).__name__}"
        if typ == "OK" and raw:
            out.append((u, raw))
            if tick:
                tick()
            continue
        failed = failed or f"UID {u}: FETCH вернул {typ}, письмо не получено"
        break
    return Fetched(out, validity, reset, failed)


# ---------------------------------------------------------------- проход

MARK_TEMP = "недоставка временно/отказ"
HARD_STATUS = re.compile(r"^5\.1\.\d+$")


def parse_dsn(msg) -> list:
    """Из частей message/delivery-status: [{recipient, action, status, diag}]. Нет Status — запись не считается разобранной."""
    out = []
    for part in msg.walk():
        st, fr = part.get("Status"), part.get("Final-Recipient") or part.get("Original-Recipient")
        if not (st or fr):
            continue
        m = re.search(r"\d\.\d{1,3}\.\d{1,3}", str(st or ""))
        out.append(dict(
            recipient=str(fr or "").split(";", 1)[-1].strip().strip("<>").lower(),
            action=str(part.get("Action", "") or "").strip().lower(),
            status=m.group(0) if m else "",
            diag=re.sub(r"\s+", " ", str(part.get("Diagnostic-Code", "") or "")).strip()[:120]))
    return out


def process(cfg: Config, fetched, stores: dict, state: dict):
    """fetched: [(папка, uid, bytes)]. Разбор писем. Возвращает (счётчики, marks, stops, новые живые).
    Сбой разбора одного письма не останавливает остальные и не теряет письмо: оно записывается как живое."""
    cnt = dict(new=0, dsn=0, auto=0, unsub=0, service=0, live=0)
    marks: dict[str, set] = {}
    stops, lives = [], []
    seen = {e.get("message_id") for e in state.get("live", []) if e.get("message_id")} | _live_ids(cfg)

    def silent(kind, folder, uid, msg, sender, why):
        log(cfg, f"{kind} [{folder} UID {uid}] от {sender or '?'}: «{str(msg.get('Subject', '') or '')[:80]}» ({why})")

    def add_live(msg, body, sender, stores, folder, in_spam, note="", uid=0):
        cnt["live"] += 1
        mid = str(msg.get("Message-ID", "") or "").strip() or f"{folder}:{uid}"   # ключ дедупликации не бывает пустым
        if mid in seen:
            return
        seen.add(mid)
        target = store_match(sender, stores)
        e = live_entry(msg, body, _domain(target) if target else (_domain(sender) if sender else ""), note)
        e["message_id"] = mid
        e["folder"] = folder
        if in_spam:
            e["webmail_hint"] += f" (папка {folder}, письмо попало в спам)"
        lives.append(e)

    for folder, uid, raw in fetched:
        _touch_lock(cfg)
        cnt["new"] += 1
        msg = email.message_from_bytes(raw, policy=policy.default)
        frm = getaddresses([str(msg.get("From", "") or "")])
        sender = frm[0][1].lower() if frm else ""
        in_spam = folder != "INBOX"
        try:
            body, rest = message_texts(msg)
            v = classify(msg, body, rest, stores)
            if in_spam and v.kind == "live" and not (store_match(sender, stores) or msg.get("In-Reply-To")):
                v = Verdict("service", "спам-папка: не ответ магазина")   # настоящий спам не поднимает тревогу
            cands = [a.lower() for a in EMAIL_RE.findall(rest + "\n" + body)]
            if v.kind == "dsn":
                recs = [r for r in parse_dsn(msg) if r["status"] and r["action"]]
                if not recs:                      # не разобрались — ничего не помечаем, но тревога
                    add_live(msg, body, sender, stores, folder, in_spam, "недоставка не разобрана", uid=uid)
                    continue
                cnt["dsn"] += 1
                for r in recs:
                    target = find_store_row(([r["recipient"]] if r["recipient"] else []) + cands, stores)
                    if not target:
                        silent("недоставка", folder, uid, msg, sender, f"адрес {r['recipient'] or '?'} не в Отправлено.xlsx")
                        continue
                    if r["action"] == "failed" and HARD_STATUS.match(r["status"]):
                        marks.setdefault(target, set()).add(MARK_DEAD)
                        stops.append((stores[target], target,
                                      f"адрес не существует ({r['status']}: {r['diag'] or 'без текста'})"))
                    else:
                        marks.setdefault(target, set()).add(MARK_TEMP)
                        silent("недоставка", folder, uid, msg, sender,
                               f"{target}: Action={r['action']} Status={r['status']} — не «не писать»")
            elif v.kind == "auto" and v.reason == "текст":
                add_live(msg, body, sender, stores, folder, in_spam, "возможно автоответ", uid=uid)
            elif v.kind == "auto":
                cnt["auto"] += 1
                target = find_store_row(cands, stores, sender)
                if target:
                    marks.setdefault(target, set()).add(MARK_AUTO)
                silent("автоответ", folder, uid, msg, sender, v.reason)
            elif v.kind == "unsub":
                cnt["unsub"] += 1
                target = store_match(sender, stores)
                stops.append((stores[target] if target else _domain(sender), sender,
                              "отписка/не писать (ответ на рассылку)"))
                silent("отписка", folder, uid, msg, sender, v.reason)
            elif v.kind == "service":
                cnt["service"] += 1
                silent("служебное", folder, uid, msg, sender, v.reason)
            else:
                add_live(msg, body, sender, stores, folder, in_spam, uid=uid)
        except Exception as e:  # noqa: BLE001  — лучше лишняя тревога, чем потерянное письмо
            log(cfg, f"[{folder} UID {uid}] ошибка разбора {type(e).__name__}: {str(e)[:100]}; записано как живое")
            try:
                add_live(msg, "", sender, stores, folder, in_spam, "ошибка разбора", uid=uid)
            except Exception:  # noqa: BLE001
                pass
    return cnt, marks, stops, lives


def run(cfg: Config | None = None, imap_factory=None, password: str | None = None) -> int:
    cfg = cfg or default_config()
    pw = password if password is not None else get_password()
    if not pw:
        log(cfg, "нет пароля, запусти Сохранить_пароль_почты.bat")
        return 2
    lock = _acquire_lock(cfg)
    if lock is None:
        log(cfg, "уже идёт другой запуск (замок state/mail_watch.lock), выхожу")
        return 4
    try:
        return _run_locked(cfg, imap_factory, pw)
    finally:
        try:
            lock.unlink()
        except OSError:
            pass


def _run_locked(cfg: Config, imap_factory, pw: str) -> int:
    t0 = time.time()
    state = _read_json(cfg.state_json, {})
    state.setdefault("live", [])
    fstate = state.setdefault("folders", {})
    if "last_uid" in state and "INBOX" not in fstate:      # состояние первой версии (одна папка)
        fstate["INBOX"] = {"last_uid": state.pop("last_uid"), "uidvalidity": state.pop("uidvalidity", None)}
    fetched, results = [], {}
    rc = 0
    M = None
    try:
        M = (imap_factory or connect)(cfg, pw)
        names = ["INBOX"]
        spam = find_spam_folder(M)
        if spam and spam.upper() != "INBOX":
            names.append(spam)
        for name in names:
            old = fstate.get(name, {})
            try:
                res = fetch_new(M, name, int(old.get("last_uid") or 0), old.get("uidvalidity"),
                                tick=lambda: _touch_lock(cfg))
            except Exception as e:  # noqa: BLE001
                if name == "INBOX":
                    raise
                log(cfg, f"папка {name} не прочитана: {type(e).__name__}: {str(e)[:100]}", secret=pw)
                rc = 1
                continue
            if res.failed:
                log(cfg, f"папка {name}: чтение оборвано ({res.failed}); курсор встанет перед этим письмом", secret=pw)
                rc = 1
            fetched += [(name, u, r) for u, r in res.got]
            results[name] = res
    except Exception as e:  # noqa: BLE001  (в тексте ошибки пароля быть не должно; вырезаем на всякий случай)
        log(cfg, f"ошибка сети/IMAP: {type(e).__name__}: {str(e)[:200]}", secret=pw)
        return 1
    finally:
        try:
            if M:
                M.logout()
        except Exception:  # noqa: BLE001
            pass
    try:
        stores = load_stores(cfg)
        cnt, marks, stops, lives = process(cfg, fetched, stores, state)
    except Exception as e:  # noqa: BLE001
        log(cfg, f"ошибка разбора: {type(e).__name__}: {str(e)[:200]}", secret=pw)
        return 3
    advance = True                 # курсор двигаем, только если всё нужное записано
    try:
        apply_marks(cfg, marks, stops)
    except Exception as e:  # noqa: BLE001  (например, xlsx открыт в Excel)
        log(cfg, f"не записан xlsx: {type(e).__name__}: {str(e)[:200]}")
        advance, rc = False, 3
    if not write_live(cfg, lives):
        advance, rc = False, 3
    else:
        state["live"].extend(lives)           # «виденным» живое считается только после записи в json
    for name, res in results.items():
        cur = fstate.setdefault(name, {})
        known = cur.get("uidvalidity")
        if res.validity and known and res.validity != known:
            cur.clear()
            cur.update(last_uid=0, uidvalidity=res.validity)      # одним действием
        elif res.validity and not known:
            cur["uidvalidity"] = res.validity
        if advance and res.got:
            cur["last_uid"] = max(int(cur.get("last_uid") or 0), max(u for u, _ in res.got))
    state["checked"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        _write_json(cfg.state_json, state)
    except OSError as e:
        log(cfg, f"не записано состояние: {e}")
        rc = 3
    log(cfg, f"новых {cnt['new']}, недоставок {cnt['dsn']}, автоответов {cnt['auto']}, отписок {cnt['unsub']}, "
             f"служебных {cnt['service']}, живых {cnt['live']} ({time.time() - t0:.1f} с)")
    return rc


def main(argv) -> int:
    if len(argv) > 1 and argv[1] == "save":
        return save_password()
    if len(argv) > 1 and argv[1] == "check":
        return check_password()
    return run()


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    sys.exit(main(sys.argv))
