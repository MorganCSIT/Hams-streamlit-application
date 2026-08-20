import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory

from planning_download import (
    PlanningFile,
    choose_latest_files,
    create_merged_planning_csv_bytes,
    get_planning_archive_folder,
    get_planning_archive_folders,
    get_planning_source_options,
    list_remote_planning_files,
    parse_planning_filename,
    planning_range_to_snapshot_dates,
    planning_files_to_rows,
    scan_all_local_planning_files,
    snapshot_date_to_planning_date,
)
from unittest.mock import patch


class PlanningDownloadTests(unittest.TestCase):
    def test_parse_planning_filename(self):
        parsed = parse_planning_filename("HAS_HAM_20260701_2330_PEPS-Visits.csv")

        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.prefix, "HAS_HAM")
        self.assertEqual(parsed.snapshot_date, "20260701")
        self.assertEqual(parsed.time, "2330")
        self.assertEqual(parsed.name, "HAS_HAM_20260701_2330_PEPS-Visits.csv")

    def test_snapshot_date_maps_to_previous_planning_date(self):
        self.assertEqual(snapshot_date_to_planning_date("20260701"), date(2026, 6, 30))

    def test_planning_range_maps_to_snapshot_dates(self):
        self.assertEqual(
            planning_range_to_snapshot_dates(date(2026, 6, 30), date(2026, 7, 2)),
            ["20260701", "20260702", "20260703"],
        )

    def test_latest_time_wins_for_duplicate_prefix_and_date(self):
        latest = choose_latest_files(
            [
                PlanningFile("HAS_HAM", "20260701", "1200", "HAS_HAM_20260701_1200_PEPS-Visits.csv"),
                PlanningFile("HAS_HAM", "20260701", "2330", "HAS_HAM_20260701_2330_PEPS-Visits.csv"),
                PlanningFile("HAS_MCT", "20260701", "0900", "HAS_MCT_20260701_0900_PEPS-Visits.csv"),
            ]
        )

        self.assertEqual(latest[("HAS_HAM", "20260701")].time, "2330")
        self.assertEqual(latest[("HAS_MCT", "20260701")].time, "0900")

    def test_source_options_include_defaults_without_local_archive(self):
        self.assertEqual(get_planning_source_options({}), ["HAS_HAM", "HAS_MCT"])

    def test_create_merged_planning_csv_bytes(self):
        with TemporaryDirectory() as tmpdir:
            first = Path(tmpdir) / "HAS_HAM_20260701_2330_PEPS-Visits.csv"
            second = Path(tmpdir) / "HAS_HAM_20260702_2330_PEPS-Visits.csv"
            first.write_text("name;value\none;1\n", encoding="utf-8")
            second.write_text("name;value\ntwo;2\n", encoding="utf-8")

            merged = create_merged_planning_csv_bytes([first, second]).decode("utf-8-sig")

        self.assertIn("source_file;planning_date;name;value", merged)
        self.assertIn("HAS_HAM_20260701_2330_PEPS-Visits.csv;2026-06-30;one;1", merged)
        self.assertIn("HAS_HAM_20260702_2330_PEPS-Visits.csv;2026-07-01;two;2", merged)

    def test_archive_folders_are_created_when_missing(self):
        with TemporaryDirectory() as tmpdir:
            expected = Path(tmpdir) / "Downloads" / "Webfleet Planning Downloads"

            with patch("planning_download.Path.home", return_value=Path(tmpdir)):
                selected = get_planning_archive_folder()
                folders = get_planning_archive_folders()

                self.assertEqual(selected, expected)
                self.assertEqual(folders, (expected,))
                self.assertTrue(expected.is_dir())

    def test_scan_all_local_planning_files_reads_both_archive_layouts(self):
        with TemporaryDirectory() as tmpdir:
            archive_folder = Path(tmpdir) / "Downloads" / "Webfleet Planning Downloads"
            archive_folder.mkdir(parents=True)
            (archive_folder / "HAS_HAM_20260701_2330_PEPS-Visits.csv").write_text("a;b\n1;2\n", encoding="utf-8")
            (archive_folder / "HAS_MCT_20260701_2330_PEPS-Visits.csv").write_text("a;b\n3;4\n", encoding="utf-8")

            with patch("planning_download.Path.home", return_value=Path(tmpdir)):
                files = scan_all_local_planning_files()

        self.assertEqual(set(files), {("HAS_HAM", "20260701"), ("HAS_MCT", "20260701")})

    def test_list_remote_planning_files_keeps_all_valid_server_files(self):
        files = list_remote_planning_files(
            [
                "README.txt",
                "HAS_HAM_20260701_2330_PEPS-Visits.csv",
                "HAS_HAM_20260701_1200_PEPS-Visits.csv",
                "HAS_MCT_20260702_2330_PEPS-Visits.csv",
            ]
        )

        self.assertEqual(len(files), 3)
        self.assertEqual(files[0].name, "HAS_MCT_20260702_2330_PEPS-Visits.csv")
        self.assertEqual(files[-1].name, "HAS_HAM_20260701_1200_PEPS-Visits.csv")

    def test_planning_files_to_rows_shows_server_inventory_fields(self):
        rows = planning_files_to_rows(
            [PlanningFile("HAS_HAM", "20260701", "2330", "HAS_HAM_20260701_2330_PEPS-Visits.csv")]
        )

        self.assertEqual(
            rows,
            [
                {
                    "prefix": "HAS_HAM",
                    "planning_date": "2026-06-30",
                    "snapshot_date": "20260701",
                    "time": "2330",
                    "file": "HAS_HAM_20260701_2330_PEPS-Visits.csv",
                }
            ],
        )


if __name__ == "__main__":
    unittest.main()
