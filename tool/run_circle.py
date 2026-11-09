"""Run quick.py over the target list of a circle: Ireland and France only, the recommended ones first.

python run_circle.py "C:/path/Цели.xlsx" --pdf-dir "C:/path/отчёты" --summary "C:/path/Сводка_прогона.xlsx"

Each target runs in its own process (quick.py, 3 s between page loads, robots.txt obeyed, no captcha or block is
bypassed). The summary is rewritten after every target, so a stopped run keeps what it has. Nothing is sent anywhere.
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

HERE = Path(__file__).parent
FIRST = ["jimmy fairly", "tediber", "bonnegueule", "le slip fran", "le petit ballon", "sheridans", "mccabes",
         "mcelhinney"]
STUDIO = "\u0441\u0442\u0443\u0434\u0438\u044f"
HEAD = ["\u2116", "\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435", "\u0421\u0430\u0439\u0442", "\u0421\u0442\u0440\u0430\u043d\u0430", "\u0422\u0438\u043f",
        "\u041f\u0440\u043e\u0432\u0435\u0440\u0435\u043d (\u0434\u0430/\u043d\u0435\u0442)", "\u041f\u0440\u0438\u0447\u0438\u043d\u0430, \u0435\u0441\u043b\u0438 \u043d\u0435\u0442",
        "\u0421\u0442\u0440\u0430\u043d\u0438\u0446 \u043f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u043e", "\u041e\u0448\u0438\u0431\u043e\u043a \u0443\u0440\u043e\u0432\u043d\u044f A",
        "\u041e\u0448\u0438\u0431\u043e\u043a \u0443\u0440\u043e\u0432\u043d\u044f AA", "\u0412\u0441\u0435\u0433\u043e \u043e\u0448\u0438\u0431\u043e\u043a", "\u041f\u0440\u043e\u0431\u043b\u0435\u043c\u0430 1",
        "\u041f\u0440\u043e\u0431\u043b\u0435\u043c\u0430 2", "\u041f\u0440\u043e\u0431\u043b\u0435\u043c\u0430 3", "PDF"]


def read_targets(path):
    ws = load_workbook(path, read_only=True, data_only=True)["\u0426\u0435\u043b\u0438"]
    rows = [r for r in ws.iter_rows(values_only=True)][1:]
    out = []
    for r in rows:
        if not r[0] or not str(r[11] or "").startswith("\u0414\u0430"):   # the letter is lawful only where it says Yes
            continue
        out.append({"n": int(r[0]), "name": str(r[1]).strip(), "site": str(r[2]).strip(), "country": r[3], "type": r[4]})
    first = [t for k in FIRST for t in out if t["name"].lower().startswith(k)]
    first += [t for t in out if t["type"] == STUDIO and t not in first]
    return first + [t for t in out if t not in first]


def safe_name(name):
    return re.sub(r'[<>:"/\|?*]', "", name).strip()


def write_summary(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "\u041f\u0440\u043e\u0433\u043e\u043d"
    ws.append(HEAD)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True)
    for col, w in zip("ABCDEFGHIJKLMNO", (5, 26, 34, 11, 10, 12, 44, 11, 11, 11, 11, 70, 70, 70, 40)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "C2"
    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("targets")
    ap.add_argument("--pdf-dir", required=True)
    ap.add_argument("--summary", required=True)
    ap.add_argument("--out", default=str(HERE / "out" / "circle1"))
    ap.add_argument("--timeout", type=int, default=420)
    args = ap.parse_args()
    out, pdf_dir = Path(args.out), Path(args.pdf_dir)
    out.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    targets = read_targets(args.targets)
    print(f"{len(targets)} targets", flush=True)
    rows = []
    for i, t in enumerate(targets, 1):
        print(f"[{i}/{len(targets)}] {t['name']} {t['site']}", flush=True)
        slug = re.sub(r"^www\.", "", re.sub(r"^https?://", "", t["site"]).split("/")[0].lower())
        folder = out / slug
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "summary.json").unlink(missing_ok=True)
        log = open(folder / "run.log", "w", encoding="utf-8")
        proc = subprocess.Popen([sys.executable, str(HERE / "quick.py"), t["site"], "--out", str(out), "--name", t["name"]],
                                stdout=log, stderr=subprocess.STDOUT)
        try:
            proc.wait(timeout=args.timeout)
        except subprocess.TimeoutExpired:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
            proc.wait()
            (folder / "summary.json").write_text(json.dumps({"error": f"timeout after {args.timeout} s"}), encoding="utf-8")
        log.close()
        try:
            s = json.loads((folder / "summary.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            s = {"error": f"no result, exit code {proc.returncode}"}
        base = [t["n"], t["name"], t["site"], t["country"], t["type"]]
        if s.get("error"):
            reason = s["error"]
            kind = "blocked" if re.search(r"HTTP 40[13]|HTTP 429|captcha|bot check", reason) else \
                "timeout" if "timeout" in reason or "no answer" in reason else "other"
            rows.append(base + ["\u043d\u0435\u0442", f"{kind}: {reason}", s.get("pages", 0), "", "", "", "", "", "", ""])
        else:
            lv = s["levels"]
            probs = [f"{f['what']} {f['who']}" for f in s["findings"]] + [""] * 3
            target_pdf = pdf_dir / f"{safe_name(t['name'])}.pdf"
            shutil.copyfile(s["pdf"], target_pdf)
            rows.append(base + ["\u0434\u0430", "", s["pages"], lv["A"]["instances"], lv["AA"]["instances"],
                                lv["total"]["instances"]] + probs[:3] + [str(target_pdf)])
        print(f"    -> {rows[-1][5]} {rows[-1][6]} pages {rows[-1][7]} A {rows[-1][8]} AA {rows[-1][9]}", flush=True)
        write_summary(rows, args.summary)
        time.sleep(3)
    ok = sum(r[5] == "\u0434\u0430" for r in rows)
    print(f"done: {ok} of {len(rows)} checked", flush=True)


if __name__ == "__main__":
    main()
