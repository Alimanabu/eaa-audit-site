"""Разовая пересборка живых данных сторожа по УЖЕ прочитанным письмам (12.11, В6).

    python tool\\rebuild_live.py              — по умолчанию --dry: показать «было / станет», ничего не записать
    python tool\\rebuild_live.py --apply      — записать (копии .bak рядом); запускает tester после приёмки
    --live-json <файл>                        — другой файл тревог (проба на копии)

Ящик только на чтение (select readonly + BODY.PEEK), курсор (last_uid) не двигается: берутся письма с UID не больше
курсора каждой папки. Разбор — тот же process()/apply_marks()/write_live(), что и у боевого запуска, но на копиях
файлов во временной папке; --apply переносит копии на место одним действием на файл. Пароль не печатается.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
from dataclasses import replace
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import mail_watch as mw  # noqa: E402


def _rows(path: Path):
    import openpyxl
    if not path.exists():
        return []
    wb = openpyxl.load_workbook(path, read_only=True)
    out = [list(r) for r in wb.active.iter_rows(values_only=True)]
    wb.close()
    return out


def _alerts(path: Path):
    d = mw._read_json(path, {})
    items = d.get("items", []) if isinstance(d, dict) else []
    return [i for i in items if isinstance(i, dict)], bool(isinstance(d, dict) and d.get("alert"))


def fetch_seen(cfg: mw.Config, state: dict, imap_factory=None, password=None):
    """[(папка, uid, bytes)] — все письма с UID не больше курсора папки (только чтение)."""
    pw = password if password is not None else mw.get_password()
    if not pw:
        raise SystemExit("нет пароля, запусти Сохранить_пароль_почты.bat")
    M = (imap_factory or mw.connect)(cfg, pw)
    out = []
    try:
        for folder, fs in (state.get("folders") or {}).items():
            last = int(fs.get("last_uid") or 0)
            if not last:
                continue
            res = mw.fetch_new(M, folder, 0, None)
            if res.failed:
                raise SystemExit(f"папка {folder}: чтение оборвано ({res.failed}), пересборку не делаю")
            out += [(folder, u, raw) for u, raw in res.got if u <= last]
    finally:
        try:
            M.logout()
        except Exception:  # noqa: BLE001
            pass
    return out


def rebuild(cfg: mw.Config, apply: bool = False, imap_factory=None, password=None, out=print) -> dict:
    state = mw._read_json(cfg.state_json, {})
    fetched = fetch_seen(cfg, state, imap_factory, password)
    if not fetched:                       # «нет данных» — не основание переписывать файлы (13.11)
        raise SystemExit("прочитано 0 писем (нет курсора или ящик пуст) — пересборку не делаю")
    with tempfile.TemporaryDirectory(prefix="eaa_rebuild_") as td:
        td = Path(td)
        tmp = replace(cfg, sent_xlsx=td / cfg.sent_xlsx.name, stop_xlsx=td / cfg.stop_xlsx.name,
                      live_json=td / "eaa_mail.json", state_json=td / "state" / "mail_watch.json", log_file=td / "rebuild.log")
        for src, dst in ((cfg.sent_xlsx, tmp.sent_xlsx), (cfg.stop_xlsx, tmp.stop_xlsx), (cfg.live_json, tmp.live_json)):
            if src.exists():
                shutil.copy2(src, dst)
        stores = mw.load_stores(tmp)
        cnt, marks, stops, lives = mw.process(tmp, fetched, stores, json.loads(json.dumps(state)))
        mw.apply_marks(tmp, marks, stops)
        if not mw.write_live(tmp, lives, stores):
            raise SystemExit("eaa_mail.json нечитаем — пересборку не делаю")
        # ---- было / станет
        a0, f0 = _alerts(cfg.live_json)
        a1, f1 = _alerts(tmp.live_json)
        s0 = {str(r[1]).lower() for r in _rows(cfg.stop_xlsx)[1:] if len(r) > 1 and r[1]}
        s1 = [r for r in _rows(tmp.stop_xlsx)[1:] if len(r) > 1 and r[1] and str(r[1]).lower() not in s0]
        x0 = {str(r[1]).lower(): r[5] for r in _rows(cfg.sent_xlsx)[1:] if len(r) > 5 and r[1]}
        x1 = {str(r[1]).lower(): r[5] for r in _rows(tmp.sent_xlsx)[1:] if len(r) > 5 and r[1]}
        res = dict(letters=len(fetched), counts=cnt,
                   alerts_before=sum(1 for i in a0 if i.get("alert")), alerts_after=sum(1 for i in a1 if i.get("alert")),
                   file_alert_before=f0, file_alert_after=f1,
                   stop_added=[(r[0], r[1], r[2]) for r in s1],
                   marks_changed={k: (x0.get(k), v) for k, v in x1.items() if x0.get(k) != v})
        out(f"Прочитано заново (только чтение): {len(fetched)} писем; разбор: {cnt}")
        out(f"eaa_mail.json: записей {len(a0)} -> {len(a1)}, с alert=true {res['alerts_before']} -> {res['alerts_after']},"
            f" alert файла {f0} -> {f1}")
        for i in a0:
            mid = i.get("message_id")
            now = next((j for j in a1 if j.get("message_id") == mid), None)
            out(f"  было  alert={i.get('alert')!s:5} {i.get('from', '')[:45]} «{str(i.get('subject', ''))[:50]}»"
                f"  -> станет {'убрана' if now is None else 'alert=' + str(now.get('alert')) + (' forwarded' if now.get('forwarded') else '')}")
        for i in a1:
            if not any(i.get("message_id") == j.get("message_id") for j in a0):
                out(f"  новая alert={i.get('alert')} {i.get('from', '')[:45]} «{str(i.get('subject', ''))[:50]}» {i.get('note', '')}")
        res["stop_confirmed"] = sorted({a for _, a, _ in stops})
        out(f"Не_писать.xlsx: по белому списку «адреса нет» и отпискам {len(res['stop_confirmed'])} адр., добавится {len(s1)}")
        for _, a, why in stops:
            out(f"  {'уже в списке' if a.lower() in s0 else '+ добавится '} {a}: {why}")
        out(f"Отправлено.xlsx, «Результат»: изменится {len(res['marks_changed'])}")
        for k, (b, a) in res["marks_changed"].items():
            out(f"  {k}: «{b}» -> «{a}»")
        if apply:
            lock = mw._acquire_lock(cfg)
            if lock is None:
                raise SystemExit("идёт запуск сторожа (замок), --apply не выполнен")
            try:
                for src, dst in ((tmp.sent_xlsx, cfg.sent_xlsx), (tmp.stop_xlsx, cfg.stop_xlsx), (tmp.live_json, cfg.live_json)):
                    if not src.exists():
                        continue
                    if dst.exists():
                        shutil.copy2(dst, dst.with_name(dst.name + ".bak"))
                    dst.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src, dst.with_name(dst.name + ".tmp"))
                    dst.with_name(dst.name + ".tmp").replace(dst)
                mw.log(cfg, f"пересборка по прочитанным письмам: тревог {res['alerts_before']} -> {res['alerts_after']}, "
                            f"в Не_писать +{len(s1)}, меток изменено {len(res['marks_changed'])}")
                out("ЗАПИСАНО (копии .bak рядом)")
            finally:
                lock.unlink()
        else:
            out("--dry: ничего не записано")
    return res


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--dry", action="store_true", default=True, help="показать было/станет (по умолчанию)")
    g.add_argument("--apply", action="store_true", help="записать")
    ap.add_argument("--live-json", default=None)
    a = ap.parse_args(argv)
    cfg = mw.default_config()
    if a.live_json:
        cfg = replace(cfg, live_json=Path(a.live_json))
    rebuild(cfg, apply=a.apply)
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    sys.exit(main())
