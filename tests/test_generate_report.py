import os
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from scripts.generate_report import generate_weekly_report, read_last_week_report


class GenerateWeeklyReportTest(unittest.TestCase):
    def setUp(self):
        self._old_output_dir = os.environ.get('WKR_OUTPUT_DIR')
        self.output_dir = tempfile.TemporaryDirectory()
        os.environ['WKR_OUTPUT_DIR'] = self.output_dir.name

    def tearDown(self):
        self.output_dir.cleanup()
        if self._old_output_dir is None:
            os.environ.pop('WKR_OUTPUT_DIR', None)
        else:
            os.environ['WKR_OUTPUT_DIR'] = self._old_output_dir

    def test_decimal_saturation_is_written_as_percent(self):
        path = generate_weekly_report({
            'period': '2026/5/26-2026/5/30',
            'name': 'Tester',
            'department': 'IT信息技术部',
            'last_week': [],
            'this_week': [],
            'next_week': [],
            'saturation': 0.9,
        })

        sheet = load_workbook(path)['员工工作周报']
        self.assertEqual(sheet['B29'].value, '90%')

    def test_output_directory_is_created_when_missing(self):
        missing_dir = Path(self.output_dir.name) / 'nested' / 'reports'
        os.environ['WKR_OUTPUT_DIR'] = str(missing_dir)

        path = generate_weekly_report({
            'period': '2026/5/26-2026/5/30',
            'last_week': [],
            'this_week': [],
            'next_week': [],
        })

        self.assertTrue(Path(path).exists())

    def test_read_last_week_report_uses_configured_output_dir(self):
        path = generate_weekly_report({
            'period': '2026/5/26-2026/5/30',
            'last_week': [],
            'this_week': [],
            'next_week': [],
        })

        self.assertEqual(read_last_week_report(), path)

    def test_iso_date_period_uses_full_range_in_filename(self):
        path = generate_weekly_report({
            'period': '2026-05-26 - 2026-05-30',
            'name': 'Tester',
            'last_week': [],
            'this_week': [],
            'next_week': [],
        })

        self.assertIn('2026-05-26至2026-05-30', Path(path).name)


if __name__ == '__main__':
    unittest.main()
