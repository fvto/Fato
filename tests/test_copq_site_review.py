"""COPQ site checks must expose empty sources and errors hidden by group totals."""
from pathlib import Path
import sys
import tempfile
import unittest

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from report_exports import check_outputs


class CopqSiteReviewTests(unittest.TestCase):
    def review(self, sources, records, mode='combine'):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            output = folder / 'workspace' / 'Output'
            output.mkdir(parents=True)
            book = openpyxl.Workbook()
            book.active.title = 'COPQ_Clean'
            book.active.append(['Site', 'Ttl cost ($)'])
            for record in records:
                book.active.append(record)
            book.save(output / 'COPQ_Clean.xlsx')
            book.close()
            return check_outputs(folder, dict(tool='copq', mode=mode, validation=sources),
                                 [dict(name='COPQ_Clean.xlsx')])

    def source(self, site, rows, cost=0, rework_rows=0):
        return dict(name=f'COPQ-{site}.xlsx', group='copq', rows=rows,
                    totals={'Ttl cost ($)': cost},
                    business_exclusions=dict(rework_rows=rework_rows, rework_cost='0'))

    def test_zero_to_zero_reconciliation_still_warns_about_empty_factory(self):
        result = self.review([self.source('VH', 0), self.source('VH2', 1, 12)], [('VH2', 12)])
        checks = {c['rule']: c for c in result['checks']}
        self.assertEqual(checks['copq.row_count']['status'], 'passed')
        self.assertEqual(checks['copq.cost']['status'], 'passed')
        self.assertEqual(result['status'], 'warning')
        warning = checks['copq.source_rows.VH']
        self.assertEqual(warning['status'], 'warning')
        self.assertIn('VH', warning['issues'][0]['title'])
        self.assertIn('Chưa đủ căn cứ', warning['issues'][0]['reason'])

    def test_swapped_sites_fail_even_when_global_totals_match(self):
        result = self.review([self.source('VH', 1, 12), self.source('JV', 1, 25)],
                             [('JV', 12), ('VH', 25)])
        checks = {c['rule']: c for c in result['checks']}
        self.assertEqual(checks['copq.cost']['status'], 'passed')
        self.assertEqual(checks['copq.site_cost.VH']['status'], 'warning')
        self.assertEqual(checks['copq.site_cost.JV']['status'], 'warning')

    def test_rework_only_is_distinct_from_empty_and_jvb_alias_is_reconciled(self):
        result = self.review([self.source('VH', 1, rework_rows=1), self.source('JVB', 1, 12)],
                             [('JV3', 12)])
        self.assertEqual(result['status'], 'passed', result)

    def test_full_report_checks_all_seven_source_factories(self):
        result = self.review([self.source('VH', 1, 12)], [('VH', 12)], mode='all')
        coverage = next(c for c in result['checks'] if c['rule'] == 'copq.site_coverage')
        self.assertEqual(coverage['status'], 'warning')
        self.assertIn('VH2', coverage['missing_sites'])


if __name__ == '__main__':
    unittest.main()
