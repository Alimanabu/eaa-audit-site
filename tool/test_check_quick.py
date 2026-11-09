"""Tests of the acceptance check for the one-page snapshot, including the negative ones.

python -m unittest test_check_quick      (run with the Python that has Pillow and reportlab)
The tests that need a real result use out/quick/fonq.nl, made by `python quick.py https://www.fonq.nl`; they are skipped
when it is not there.
"""
import json
import shutil
import tempfile
import unittest
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

import check_quick as C

GOOD_TEXT = ("4 pages checked 25 violations found on 20 distinct elements 4 of 50 WCAG 2.1 A/AA criteria not passed "
             "The full automated audit (€390) checks up to 10 pages. Details: https://eaa-audit.net")
EXPECTED = {"pages": 4, "violations": 25, "elements": 20, "criteria": 4}
REAL = Path(__file__).parent / "out" / "quick" / "fonq.nl"


class TextChecks(unittest.TestCase):
    def test_matching_numbers_pass(self):
        self.assertEqual(C.text_problems(GOOD_TEXT, EXPECTED), [])

    def test_wrong_violation_count_fails(self):
        bad = C.text_problems(GOOD_TEXT, {**EXPECTED, "violations": 24})
        self.assertEqual(len(bad), 1)
        self.assertIn("PDF says 25, axe results say 24", bad[0])

    def test_second_source_zeroed_fails(self):
        """The axe results are emptied: the numbers in the PDF no longer agree, so the status is not ok."""
        raw = C.raw_numbers({"pages": []})
        self.assertEqual((raw["pages"], raw["violations"], raw["elements"], raw["criteria"]), (0, 0, 0, 0))
        self.assertEqual(len(C.text_problems(GOOD_TEXT, raw)), 4)

    def test_forbidden_words_fail(self):
        for word in ("a penalty of 5", "penalties apply", "a fine is possible", "Fines vary"):
            self.assertTrue(any("forbidden word" in b for b in C.text_problems(GOOD_TEXT + " " + word, EXPECTED)), word)
        self.assertEqual(C.text_problems(GOOD_TEXT + " a definite refinement", EXPECTED), [])  # not the word "fine"

    def test_other_euro_amounts_fail(self):
        self.assertTrue(C.text_problems(GOOD_TEXT + " up to €5,000", EXPECTED))
        self.assertTrue(C.text_problems(GOOD_TEXT.replace("€390", "390 EUR"), EXPECTED))  # the price must be there

    def test_missing_link_fails(self):
        self.assertTrue(C.text_problems(GOOD_TEXT.replace("https://eaa-audit.net", ""), EXPECTED))

    def test_raw_numbers_count_nodes_once_per_element(self):
        rule = lambda i, targets: {"id": i, "criteria": ["1.4.3"], "nodeCount": len(targets),
                                   "nodes": [{"target": [t]} for t in targets]}
        raw = C.raw_numbers({"pages": [{"n": 1, "violations": [rule("a", ["x", "y"]), rule("b", ["x"])]},
                                       {"n": 2, "violations": [rule("a", ["x"])]}, {"n": 3, "error": "x"}]})
        self.assertEqual((raw["pages"], raw["violations"], raw["elements"], raw["criteria"]), (2, 4, 2, 1))


@unittest.skipUnless((REAL / "summary.json").exists(), "run quick.py on fonq.nl first")
class RealFolder(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.folder = self.tmp / "fonq.nl"
        shutil.copytree(REAL, self.folder, ignore=shutil.ignore_patterns("preview*"))

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_untouched_folder_passes(self):
        self.assertEqual(C.check_folder(self.folder), [])

    def test_changed_axe_results_fail(self):
        path = self.folder / "results.json"
        data = json.loads(path.read_text(encoding="utf-8"))
        rule = next(v for p in data["pages"] for v in p.get("violations", []))
        rule["nodeCount"] += 1
        path.write_text(json.dumps(data), encoding="utf-8")
        self.assertTrue(any("violations: PDF says" in b for b in C.check_folder(self.folder)))

    def test_two_page_pdf_fails(self):
        pdf = next(self.folder.glob("*.pdf"))
        c = canvas.Canvas(str(pdf), pagesize=A4)
        c.drawString(72, 700, "page one")
        c.showPage()
        c.drawString(72, 700, "page two")
        c.save()
        self.assertTrue(any("must be 1" in b for b in C.check_folder(self.folder)))

    def test_picture_without_frame_fails(self):
        from PIL import Image
        picture = next((self.folder / "pictures").glob("*.jpg"))
        Image.new("RGB", (700, 300), "white").save(picture)
        self.assertTrue(any("no red frame" in b for b in C.check_folder(self.folder)))


if __name__ == "__main__":
    unittest.main()
