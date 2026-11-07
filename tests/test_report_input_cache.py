import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

import report_input_cache as cache


class InputCacheTests(unittest.TestCase):
    def test_reuses_identical_bytes_and_rechecks_changes_even_with_same_timestamp(self):
        with TemporaryDirectory() as folder, patch.object(cache, 'APP_ROOT', Path(folder)):
            path = Path(folder) / 'input.xlsx'
            path.write_bytes(b'original')
            check = Mock(return_value=True)
            self.assertTrue(cache.eligible(path, 'hfpa', ('hfpa', '2026-09'), check))
            renamed = path.with_name('renamed.xlsx')
            renamed.write_bytes(path.read_bytes())
            self.assertTrue(cache.eligible(renamed, 'hfpa', ('hfpa', '2026-09'), check))
            self.assertEqual(check.call_count, 1)
            stat = path.stat()
            path.write_bytes(b'modified')
            os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns))
            check.return_value = False
            self.assertFalse(cache.eligible(path, 'hfpa', ('hfpa', '2026-09'), check))
            self.assertEqual(check.call_count, 2)
            cache.eligible(path, 'hfpa', ('hfpa', '2026-10'), check)
            self.assertEqual(check.call_count, 3)

    def test_corrupt_summary_and_unwritable_cache_fall_back_to_validation(self):
        with TemporaryDirectory() as folder, patch.object(cache, 'APP_ROOT', Path(folder)):
            path = Path(folder) / 'input.xlsx'
            path.write_bytes(b'original')
            check = Mock(return_value=True)
            cache.eligible(path, 'copq', ('ftt',), check)
            for saved in Path(folder).rglob('*.json'):
                saved.write_text('{broken', encoding='utf-8')
            with patch('sfc_report_tools.write_json', side_effect=PermissionError):
                self.assertTrue(cache.eligible(path, 'copq', ('ftt',), check))
            self.assertEqual(check.call_count, 2)
