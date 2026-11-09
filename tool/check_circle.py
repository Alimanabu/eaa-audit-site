"""Check the result of run_circle.py: the summary workbook against the PDFs.

python check_circle.py "Сводка_прогона.xlsx"
For every row marked as checked: the PDF exists and has 1 page; its text is English (no Cyrillic) and has no empty
field or service word; the three problems of the summary are in the PDF word for word; the numbers of the summary
(A + AA = total) equal the number in the PDF. Every row not checked has a reason. Exit code 1 on any failure.
"""
import re
import subprocess
import sys

from openpyxl import load_workbook

SERVICE = re.compile(r"\b(None|null|nan|TODO|TBD|undefined|info page)\b|\{\{|\[\]", re.I)
CYR = re.compile("[\u0400-\u04ff]")


def run(*cmd):
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace").stdout


def main():
    rows = [r for r in load_workbook(sys.argv[1]).active.iter_rows(min_row=2, values_only=True) if r[0]]
    ok = bad = 0
    for r in rows:
        n, name, site, checked, reason, pages, a, aa, total, p1, p2, p3, pdf = r[0], r[1], r[2], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[13], r[14]
        problems = []
        if checked == "\u0434\u0430":
            text = run("pdftotext", "-enc", "UTF-8", pdf, "-")
            flat = re.sub(r"\s+", " ", text)
            if not re.search(r"Pages:\s+1\b", run("pdfinfo", pdf)):
                problems.append("PDF is not 1 page")
            if CYR.search(text):
                problems.append("Cyrillic in the PDF")
            if SERVICE.search(text):
                problems.append(f"service word: {SERVICE.search(text).group(0)}")
            m = re.search(r"(\d+) violations found on", flat)
            if not m or int(m.group(1)) != (a or 0) + (aa or 0) or total != (a or 0) + (aa or 0):
                problems.append(f"violations: PDF {m.group(1) if m else '?'}, summary A {a} + AA {aa}, total {total}")
            if not re.search(rf"\b{pages} pages checked", flat):
                problems.append(f"pages: summary {pages} not in the PDF")
            for k, pr in enumerate((p1, p2, p3), 1):
                if not pr:
                    problems.append(f"problem {k} is empty")
                elif pr.startswith(("No further problem type", "The automated rules found no failures")):
                    pass  # the filler for a shop with fewer than three problem types
                elif re.sub(r"\s+", " ", pr) not in flat.replace("What is wrong: ", "").replace("Who it affects: ", ""):
                    # the PDF prints the two sentences with labels in between: compare the sentences one by one
                    parts = re.split(r"(?<=\.)\s(?=[A-Z])", pr, maxsplit=1)
                    if not all(part.strip().replace("\n", " ")[:60] in flat for part in parts):
                        problems.append(f"problem {k} differs from the PDF")
            ok += not problems
        else:
            if not reason:
                problems.append("not checked, no reason")
        bad += bool(problems)
        print(f"{'OK  ' if not problems else 'FAIL'} {n:>2} {name}: {'checked' if checked == chr(0x434)+chr(0x430) else 'NOT checked: ' + str(reason)[:70]}")
        for p in problems:
            print("      -", p)
    checked_n = sum(r[5] == "\u0434\u0430" for r in rows)
    print(f"{checked_n} of {len(rows)} checked; {bad} rows with problems")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
