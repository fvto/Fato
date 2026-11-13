from io import BytesIO
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
from urllib.error import HTTPError, URLError
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import sfc_qms as qms
import sfc_qms_bulk as bulk


class BulkPdfTests(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.root = Path(folder.name)
        root_patch = patch.object(bulk, 'DOWNLOAD_ROOT', self.root)
        jobs_patch = patch.object(bulk, 'JOBS', {})
        root_patch.start()
        jobs_patch.start()
        self.addCleanup(root_patch.stop)
        self.addCleanup(jobs_patch.stop)
        self.entries = {f'{index:032x}': dict(id=f'{index:032x}', source='manual', code=f'S{index:04}', name='Shoe', language='')
                        for index in range(25)}
        self.catalog = patch.object(qms, 'read_json', wraps=qms.read_json)

    def seed(self, ids=None):
        key = 'a' * 32
        ids = list(self.entries) if ids is None else ids
        job = dict(id=key, plant='VH', source='manual', ids=ids, total=len(ids), completed=0, status='queued', current='',
                   error=None, filename='PDFs.zip', bytes=0)
        bulk.JOBS[key] = job
        bulk.save(job)
        return key

    def test_all_25_files_and_colliding_names_use_one_isolated_session(self):
        key = self.seed()
        with patch.object(qms, 'DocumentSession') as session, \
             patch.object(qms, 'download_pdf', side_effect=lambda *a, **kw: (BytesIO(b'%PDF-1.7\nTest\n%%EOF'), 'same.pdf')) as download:
            bulk.work(key, self.entries)
        result = bulk.snapshot(key)
        self.assertEqual((result['status'], result['completed'], result['total']), ('done', 25, 25))
        self.assertEqual(download.call_count, 25)
        session.return_value.prepare_catalog.assert_called_once()
        self.assertTrue(all(call.kwargs['session'] is session.return_value for call in download.call_args_list))
        with ZipFile(bulk.archive_file(key)) as archive:
            self.assertEqual(len(archive.namelist()), 25)
            self.assertEqual(len(set(archive.namelist())), 25)
            self.assertIsNone(archive.testzip())
        self.assertFalse((self.root / key / 'archive.partial').exists())

    def test_connection_failure_preserves_files_and_labels_partial_zip(self):
        key = self.seed(ids=list(self.entries)[:3])
        with patch.object(qms, 'DocumentSession'), patch.object(qms, 'download_pdf', side_effect=[
            (BytesIO(b'%PDF-1.7\nTest\n%%EOF'), 'first.pdf'), RuntimeError('SFC hết đăng nhập')]) as download:
            bulk.work(key, self.entries)
        result = bulk.snapshot(key)
        self.assertEqual((result['status'], result['completed']), ('partial', 1))
        self.assertEqual(download.call_count, 2)
        self.assertIn('url', result)
        self.assertIn('THIEU_2_PDF', result['filename'])
        self.assertFalse((self.root / key / 'archive.partial').exists())
        self.assertTrue((self.root / key / 'pdfs' / (next(iter(self.entries)) + '.pdf')).is_file())
        with ZipFile(bulk.archive_file(key)) as archive:
            self.assertIn('PDF_CHUA_TAI_DUOC.csv', archive.namelist())

    def test_missing_pdf_does_not_block_later_files_and_resume_only_fetches_missing(self):
        ids = list(self.entries)[:3]
        key = self.seed(ids=ids)
        good = lambda name: (BytesIO(b'%PDF-1.7\nTest\n%%EOF'), name)
        with patch.object(qms, 'DocumentSession'), patch.object(qms, 'download_pdf', side_effect=[
            good('one.pdf'), HTTPError('https://sfc.chingluh.com/missing', 404, 'Not Found', {}, None), good('three.pdf')]) as download:
            bulk.work(key, self.entries)
        result = bulk.snapshot(key)
        self.assertEqual((result['status'], result['completed'], result['missing_count']), ('partial', 2, 1))
        self.assertEqual(download.call_count, 3)
        self.assertTrue(result['resumable'])
        self.assertEqual(result['failed'][0]['kind'], 'missing')
        with ZipFile(bulk.archive_file(key)) as archive:
            report = archive.read('PDF_CHUA_TAI_DUOC.csv').decode('utf-8-sig')
            self.assertIn('S0001', report)
            self.assertIn('404', report)
            self.assertEqual(sum(n.endswith('.pdf') for n in archive.namelist()), 2)
        catalog = self.root / 'catalog.json'
        qms.atomic_json(catalog, {'VH': {'documents': list(self.entries.values())}})
        bulk.JOBS.clear()
        with patch.object(qms, 'CATALOG', catalog), patch.object(bulk.threading, 'Thread'):
            self.assertEqual(bulk.resume(key)['status'], 'queued')
        with patch.object(qms, 'DocumentSession'), patch.object(qms, 'download_pdf', side_effect=lambda *a, **kw: good('two.pdf')) as download:
            bulk.work(key, self.entries)
        self.assertEqual(bulk.snapshot(key)['status'], 'done')
        self.assertEqual(bulk.snapshot(key)['cached'], 2)
        download.assert_called_once()
        self.assertEqual(download.call_args.args[1], ids[1])
        with ZipFile(bulk.archive_file(key)) as archive:
            self.assertEqual(len(archive.namelist()), 3)
            self.assertNotIn('PDF_CHUA_TAI_DUOC.csv', archive.namelist())

    def test_errors_distinguish_missing_source_network_and_disk(self):
        import errno
        self.assertEqual(qms.download_error(HTTPError('https://sfc.chingluh.com/file', 404, 'Not Found', {}, None))['kind'], 'missing')
        self.assertEqual(qms.download_error(URLError('timed out'))['kind'], 'connection')
        self.assertEqual(qms.download_error(OSError(errno.ENOSPC, 'full'))['kind'], 'disk_full')
        self.assertEqual(qms.download_error(PermissionError('locked'))['kind'], 'write_permission')

    def test_start_validates_catalog_and_deduplicates_request_and_blocks_overlap(self):
        catalog = self.root / 'catalog.json'
        qms.atomic_json(catalog, {'VH': {'documents': list(self.entries.values())}})
        with patch.object(qms, 'CATALOG', catalog), patch.object(bulk.threading, 'Thread') as thread:
            job = bulk.start('VH', list(self.entries) + [next(iter(self.entries))], 'manual')
            self.assertEqual(job['total'], 25)
            self.assertEqual(job['source'], 'manual')
            self.assertTrue(job['filename'].startswith('QA_Manual_'))
            thread.return_value.start.assert_called_once()
            with self.assertRaises(RuntimeError): bulk.start('VH', list(self.entries), 'manual')
            with self.assertRaises(ValueError): bulk.start('VH', ['f' * 32], 'manual')
            with self.assertRaises(ValueError): bulk.start('VH', [], 'manual')
            with self.assertRaises(ValueError): bulk.start('VH', ['../bad'], 'manual')

    def test_manual_and_document_cannot_be_combined_or_mislabeled(self):
        catalog = self.root / 'catalog.json'
        entries = list(self.entries.values())
        entries[-1] = dict(entries[-1], source='document')
        qms.atomic_json(catalog, {'VH': {'documents': entries}})
        with patch.object(qms, 'CATALOG', catalog), patch.object(bulk.threading, 'Thread') as thread:
            with self.assertRaises(ValueError): bulk.start('VH', list(self.entries))
            with self.assertRaises(ValueError): bulk.start('VH', list(self.entries), 'manual')
            with self.assertRaises(ValueError): bulk.start('VH', [entries[0]['id']], 'document')
            thread.assert_not_called()
            job = bulk.start('VH', [entries[-1]['id']], 'document')
            self.assertEqual(job['source'], 'document')
            self.assertTrue(job['filename'].startswith('QA_Document_'))

    def test_completed_zip_survives_reload_and_interrupted_job_is_reported(self):
        key = self.seed(ids=list(self.entries)[:1])
        with patch.object(qms, 'DocumentSession'), patch.object(qms, 'download_pdf', return_value=(BytesIO(b'%PDF-1.7\nTest\n%%EOF'), 'one.pdf')):
            bulk.work(key, self.entries)
        bulk.JOBS.clear()
        self.assertEqual(bulk.snapshot(key)['status'], 'done')
        self.assertTrue(bulk.archive_file(key).is_file())
        job = qms.read_json(self.root / key / 'job.json')
        job['status'] = 'running'
        qms.atomic_json(self.root / key / 'job.json', job)
        self.assertEqual(bulk.snapshot(key)['status'], 'error')
        with self.assertRaises(FileNotFoundError): bulk.snapshot('../bad')


if __name__ == '__main__':
    unittest.main()
