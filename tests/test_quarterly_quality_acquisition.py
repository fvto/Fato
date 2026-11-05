from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import sfc_quarterly_quality_acquisition as acquisition


class QuarterlyQualityAcquisitionTests(unittest.TestCase):
    def setUp(self):
        self.source = patch('report_tools.quarterly_quality.tableau_source.default_path', return_value=None)
        self.source.start()
        self.addCleanup(self.source.stop)

    def test_plan_requests_only_the_selected_quarter(self):
        with patch.object(acquisition, "local_today", return_value=date(2026, 11, 5)):
            for cutoff, count in (("2026-09-30", 12), ("2026-10-31", 24), ("2026-11-05", 36)):
                with self.subTest(cutoff=cutoff):
                    plan = acquisition.plan("FY27-Q2", cutoff)
                    self.assertEqual(plan["total"], count)
                    self.assertEqual(len(plan["missing"]), count)
                    self.assertTrue(all("2026-09-01" <= slot["date_from"] <= slot["date_to"] <= cutoff
                                        for slot in plan["missing"]))
                    self.assertEqual({slot["plant"] for slot in plan["missing"]}, {"VH", "VH2", "JV", "JV2"})
                    self.assertEqual({slot["group"] for slot in plan["missing"]}, {"hfpa", "qastation", "ftt"})
                    self.assertNotIn("prior_date_from", plan)
                    self.assertNotIn("prior_date_to", plan)

    def test_plan_handles_a_quarter_crossing_the_calendar_year(self):
        with patch.object(acquisition, "local_today", return_value=date(2027, 2, 5)):
            plan = acquisition.plan("FY27-Q3", "2027-02-05")
        self.assertEqual(plan["date_from"], "2026-12-01")
        self.assertEqual(plan["total"], 36)
        self.assertEqual({slot["date_from"][:7] for slot in plan["missing"]}, {"2026-12", "2027-01", "2027-02"})

    def test_resuming_an_older_run_discards_prior_quarter_downloads(self):
        prior = dict(group="hfpa", plant="VH", date_from="2026-06-01", date_to="2026-06-30")
        current = dict(group="hfpa", plant="VH", date_from="2026-09-01", date_to="2026-09-30")
        run = dict(period="FY27-Q2", inputs=[], acquisition=dict(
            date_from="2026-09-01", date_to="2026-09-30", total=24, reused=[], missing=[prior, current]))
        with TemporaryDirectory() as folder, patch.object(acquisition.reports, "EXPORT_ROOT", Path(folder)), \
                patch("sfc_report_acquisition.acquire_month") as download:
            acquisition.acquire(Path(folder) / "run", run, lambda *_: None)
        self.assertEqual(run["acquisition"]["missing"], [current])
        self.assertEqual(run["acquisition"]["total"], 12)
        download.assert_called_once()


if __name__ == "__main__":
    unittest.main()
