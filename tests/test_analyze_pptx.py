"""Acceptance checks against the real 22-slide editable deck."""

import json
import tempfile
import unittest
from collections import Counter
from pathlib import Path

from analyze_pptx import extract, flatten, validate_review
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


ROOT = Path(__file__).resolve().parents[1]
DECK = ROOT / "Paper-Presentation-Tips-editable-22-v2.pptx"


class SyntheticExtraction(unittest.TestCase):
    def test_native_table_chart_and_notes_are_separated(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "mini.pptx"
            pres = Presentation()
            slide = pres.slides.add_slide(pres.slide_layouts[6])
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = "实验结论"
            table = slide.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(4), Inches(1)).table
            table.cell(0, 0).text = "指标"
            table.cell(0, 1).text = "结果"
            table.cell(1, 0).text = "准确率"
            table.cell(1, 1).text = "84"
            data = CategoryChartData()
            data.categories = ["A", "B"]
            data.add_series("模型", [84, 96])
            slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6), Inches(1),
                                   Inches(4), Inches(3), data)
            slide.notes_slide.notes_text_frame.text = "只供讲者看"
            pres.save(path)
            item = extract(path)["slides"][0]
            self.assertIn("84", item["visible_text"])
            self.assertNotIn("只供讲者看", item["visible_text"])
            self.assertIn("只供讲者看", item["notes_not_visible"])
            chart = next(obj["chart"] for obj in item["objects"] if "chart" in obj)
            self.assertEqual([84.0, 96.0], chart["series"][0]["values"])


@unittest.skipUnless(DECK.exists(), "v2 acceptance PPTX is absent")
class RealDeckAcceptance(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.inventory = extract(DECK)
        cls.slides = cls.inventory["slides"]

    def test_all_slides_and_object_kinds(self):
        self.assertEqual(22, self.inventory["slide_count"])
        counts = Counter()
        for slide in self.slides:
            for obj in flatten(slide["objects"]):
                for kind in ("table", "chart", "picture"):
                    if kind in obj:
                        counts[kind] += 1
        self.assertEqual({"table": 6, "chart": 3, "picture": 19}, dict(counts))

    def test_table_cells_and_chart_values_are_bound_to_correct_slides(self):
        table = next(obj["table"] for obj in self.slides[17]["objects"] if "table" in obj)
        self.assertEqual((7, 4), (table["row_count"], table["column_count"]))
        self.assertTrue(any(cell["text"] == "237.9" for row in table["rows"] for cell in row))
        charts = [obj["chart"] for obj in self.slides[19]["objects"] if "chart" in obj]
        self.assertEqual(2, len(charts))
        self.assertEqual([51.5, 65.9], charts[0]["series"][0]["values"])
        self.assertEqual([253.0, 263.3], charts[1]["series"][1]["values"])
        closing_chart = next(obj["chart"] for obj in self.slides[21]["objects"] if "chart" in obj)
        self.assertEqual([84.0, 71.0, 96.0], closing_chart["series"][0]["values"])

    def test_visibility_and_review_binding(self):
        self.assertIn("机器翻译", "".join(self.slides[1]["visible_text"]))
        self.assertTrue(all(slide["render"]["status"] == "ok"
                            for slide in json.loads((ROOT / "audit-v2/inventory.json").read_text(encoding="utf-8"))["slides"]))
        review = json.loads((ROOT / "validation/v2_review.json").read_text(encoding="utf-8"))
        rendered_inventory = json.loads((ROOT / "audit-v2/inventory.json").read_text(encoding="utf-8"))
        validate_review(review, rendered_inventory)
        self.assertTrue(any(slide["status"] == "无需修改" for slide in review["slides"]))
        self.assertTrue(any(slide["status"] == "有确定问题" for slide in review["slides"]))


if __name__ == "__main__":
    unittest.main()
