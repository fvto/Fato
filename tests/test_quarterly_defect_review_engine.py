import tempfile
import unittest
from pathlib import Path

import openpyxl

from report_tools.quarterly_defect_review.analysis import analyze_metric, fmt_pct
from report_tools.quarterly_defect_review.area_mapping import (
    build_area_catalog,
    classify_areas,
    normalize_line_code,
)
from report_tools.quarterly_defect_review.defect_mapping import load_defect_mapping, map_defects
from report_tools.quarterly_defect_review.nike_inputs import normalize_nike_source


def _run_memory(metric, rows, catalog_entries=None, mapping=None):
    norm = normalize_nike_source(
        [{'source_name': f'{metric.lower()}_mem', 'report_mode': 'NIKE', 'rows': rows}],
        metric=metric,
        mapping=mapping,
    )
    catalog = build_area_catalog(entries=catalog_entries or [])
    base_recs, facts, area_reviews = classify_areas(norm['base_records'], norm['defect_facts'], catalog)
    reviews = list(norm['review_items']) + list(area_reviews)
    result = analyze_metric(
        metric=metric,
        base_records=base_recs,
        defect_facts=facts,
        review_items=reviews,
        manifests=norm['manifests'],
        detail_support_level=norm['detail_support_level'],
        catalog_version=catalog['catalog_version'],
        mapping_version=norm['mapping_version'],
        mapping=mapping,
    )
    return result, norm, catalog


class QuarterlyDefectReviewEngineTests(unittest.TestCase):
    def test_01_fixture_bonding_1000_exact_math_and_denominators(self):
        """Section 9 & 14: Exact 1,000 Bonding fixture verifying 40/35/25, 65/25/7/3, 75.0%, 46.2%."""
        catalog_entries = [
            {'factory': 'VH', 'plant': 'N14', 'line_raw': 'Line 1', 'source_group': 'Factory Sandals'},
        ]
        # VH (S2 Line 1), JV (JV1 Line 1), VB (N14 Line 1 matching Factory Sandals)
        rows = [
            # Lỗi A: VH=300, JV=200, VB=150 -> CLG=650
            {'audit_id': 'A-VH', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 1000, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Bond gap / Rat hole', 'detail_defect': 'Lỗi A', 'issue_quantity': 300},
            {'audit_id': 'A-JV', 'date': '2026-06-05', 'factory': 'JV', 'plant': 'P1', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 1000, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Bond gap / Rat hole', 'detail_defect': 'Lỗi A', 'issue_quantity': 200},
            {'audit_id': 'A-VB', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'N14', 'line': '1', 'model': 'SLIDE 1', 'inspected_qty': 800, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Bond gap / Rat hole', 'detail_defect': 'Lỗi A', 'issue_quantity': 150},
            # Lỗi B: VH=70, JV=110, VB=70 -> CLG=250
            {'audit_id': 'B-VH', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M2', 'inspected_qty': 500, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Delamination', 'detail_defect': 'Lỗi B', 'issue_quantity': 70},
            {'audit_id': 'B-JV', 'date': '2026-06-05', 'factory': 'JV2', 'plant': 'P2', 'line': 'Line 2', 'model': 'M2', 'inspected_qty': 500, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Delamination', 'detail_defect': 'Lỗi B', 'issue_quantity': 110},
            {'audit_id': 'B-VB', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'N14', 'line': 'Line 1', 'model': 'SLIDE 1', 'inspected_qty': 500, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Delamination', 'detail_defect': 'Lỗi B', 'issue_quantity': 70},
            # Lỗi C: VH=20, JV=30, VB=20 -> CLG=70
            {'audit_id': 'C-VH', 'date': '2026-06-06', 'factory': 'VH2', 'plant': 'S3', 'line': 'Line 2', 'model': 'M3', 'inspected_qty': 200, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Bond gap / Rat hole', 'detail_defect': 'Lỗi C', 'issue_quantity': 20},
            {'audit_id': 'C-JV', 'date': '2026-06-06', 'factory': 'JV', 'plant': 'P1', 'line': 'Line 3', 'model': 'M3', 'inspected_qty': 200, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Bond gap / Rat hole', 'detail_defect': 'Lỗi C', 'issue_quantity': 30},
            {'audit_id': 'C-VB', 'date': '2026-06-06', 'factory': 'VH', 'plant': 'N14', 'line': 'Line1', 'model': 'SLIDE 2', 'inspected_qty': 200, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Bond gap / Rat hole', 'detail_defect': 'Lỗi C', 'issue_quantity': 20},
            # Lỗi D: VH=10, JV=10, VB=10 -> CLG=30
            {'audit_id': 'D-VH', 'date': '2026-06-06', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 4', 'model': 'M4', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Delamination', 'detail_defect': 'Lỗi D', 'issue_quantity': 10},
            {'audit_id': 'D-JV', 'date': '2026-06-06', 'factory': 'JV', 'plant': 'P1', 'line': 'Line 4', 'model': 'M4', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Delamination', 'detail_defect': 'Lỗi D', 'issue_quantity': 10},
            {'audit_id': 'D-VB', 'date': '2026-06-06', 'factory': 'VH', 'plant': 'N14', 'line': '1', 'model': 'SLIDE 3', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Delamination', 'detail_defect': 'Lỗi D', 'issue_quantity': 10},
        ]
        res, _, _ = _run_memory('FTT', rows, catalog_entries)
        bonding = res['top_groups'][0]
        self.assertEqual(bonding['group_id'], 'bond_gap_bonding')
        self.assertEqual(bonding['clg_qty'], 1000.0)
        self.assertEqual(bonding['vh_qty'], 400.0)
        self.assertEqual(bonding['jv_qty'], 350.0)
        self.assertEqual(bonding['vb_qty'], 250.0)
        self.assertEqual(fmt_pct(bonding['vh_contrib_pct']), '40.0%')
        self.assertEqual(fmt_pct(bonding['jv_contrib_pct']), '35.0%')
        self.assertEqual(fmt_pct(bonding['vb_contrib_pct']), '25.0%')

        clg_cell = res['details_by_group']['bond_gap_bonding']['CLG']
        self.assertEqual([d['detail_label'] for d in clg_cell['top_details']], ['Lỗi A', 'Lỗi B', 'Lỗi C'])
        self.assertEqual([d['qty'] for d in clg_cell['top_details']], [650.0, 250.0, 70.0])
        self.assertEqual([fmt_pct(d['share_of_group_in_scope_pct']) for d in clg_cell['top_details']], ['65.0%', '25.0%', '7.0%'])
        self.assertEqual(clg_cell['other_mapped_details_qty'], 30.0)
        self.assertEqual(fmt_pct(clg_cell['other_mapped_details_pct']), '3.0%')

        vh_cell = res['details_by_group']['bond_gap_bonding']['VH']
        vh_top1 = vh_cell['top_details'][0]
        self.assertEqual(vh_top1['detail_label'], 'Lỗi A')
        self.assertEqual(vh_top1['qty'], 300.0)
        # Denominator 1: share of Bonding in VH = 300/400 = 75.0%
        self.assertEqual(fmt_pct(vh_top1['share_of_group_in_scope_pct']), '75.0%')
        # Denominator 2: VH contribution to Lỗi A CLG = 300/650 = 46.2% (NEVER swapped with 75.0%!)
        self.assertEqual(fmt_pct(vh_top1['area_contribution_to_detail_pct']), '46.2%')

        dd_a = res['drilldown_by_group']['bond_gap_bonding']['details'][0]
        self.assertEqual(dd_a['detail_label'], 'Lỗi A')
        self.assertEqual(fmt_pct(dd_a['share_of_clg_group_pct']), '65.0%')
        self.assertEqual(fmt_pct(dd_a['vh_contrib_to_detail_pct']), '46.2%')

        # DR remains explicitly unconfirmed
        self.assertEqual(res['dr_config']['status'], 'unconfirmed')
        self.assertIsNone(res['dr_config']['formula_name'])

    def test_02_two_observed_defects_never_fabricated_to_three(self):
        rows = [
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'detail_defect': 'Outsole bond gap', 'issue_quantity': 60},
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'detail_defect': 'Toe delamination', 'issue_quantity': 40},
        ]
        res, _, _ = _run_memory('FTT', rows)
        top_details = res['details_by_group']['bond_gap_bonding']['CLG']['top_details']
        self.assertEqual(len(top_details), 2)
        self.assertEqual([d['detail_label'] for d in top_details], ['Outsole bond gap', 'Toe delamination'])

    def test_03_hfpa_group_only_reports_missing_detail_and_never_uses_ftt_to_split(self):
        hfpa_rows = [
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'NIKE PEGASUS', 'inspected_qty': 200, 'standard_group': 'Bond gap / Bonding', 'issue_quantity': 120},
        ]
        res_hfpa, _, _ = _run_memory('HFPA', hfpa_rows)
        cell = res_hfpa['details_by_group']['bond_gap_bonding']['CLG']
        self.assertEqual(cell['top_details'], [])
        self.assertEqual(cell['cell_status'], 'missing_detail')
        self.assertIn('Thiếu dữ liệu chi tiết', cell['status_note'])
        # Top Models still populated in parallel even when sub-detail is missing!
        self.assertEqual(res_hfpa['models_by_group']['bond_gap_bonding']['CLG']['top_models'][0]['model_label'], 'NIKE PEGASUS')
        self.assertEqual(res_hfpa['models_by_group']['bond_gap_bonding']['CLG']['top_models'][0]['qty'], 120.0)

    def test_04_hfpa_vb_classified_from_ftt_catalog_preserving_hfpa_qty(self):
        catalog_entries = [{'factory': 'VH', 'plant': 'N14', 'line_raw': 'Line 1', 'source_group': 'Factory Sandals'}]
        hfpa_rows = [
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'N14', 'line': '1', 'model': 'SLIDE X', 'inspected_qty': 100, 'standard_group': 'Cement', 'issue_quantity': 45},
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'SHOE Y', 'inspected_qty': 100, 'standard_group': 'Cement', 'issue_quantity': 55},
        ]
        res, _, _ = _run_memory('HFPA', hfpa_rows, catalog_entries=catalog_entries)
        cement = res['top_groups'][0]
        self.assertEqual(cement['vb_qty'], 45.0)
        self.assertEqual(cement['vh_qty'], 55.0)
        self.assertEqual(cement['clg_qty'], 100.0)

    def test_05_sandals_move_from_vh_to_vb_without_double_counting_clg(self):
        rows = [
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'N14', 'line': 'Line 2', 'model': 'SLIDE A', 'inspected_qty': 200, 'standard_group': 'Cement', 'detail_defect': 'Over cement', 'issue_quantity': 80},
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 2', 'model': 'SHOE B', 'inspected_qty': 200, 'standard_group': 'Cement', 'detail_defect': 'Over cement', 'issue_quantity': 120},
        ]
        res_no_vb, _, _ = _run_memory('FTT', rows, catalog_entries=[])
        res_with_vb, _, _ = _run_memory('FTT', rows, catalog_entries=[{'factory': 'VH', 'plant': 'N14', 'line_raw': 'Line 2', 'source_group': 'Factory Sandals'}])

        g0 = res_no_vb['top_groups'][0]
        g1 = res_with_vb['top_groups'][0]
        self.assertEqual((g0['vh_qty'], g0['vb_qty'], g0['clg_qty']), (200.0, 0.0, 200.0))
        self.assertEqual((g1['vh_qty'], g1['vb_qty'], g1['clg_qty']), (120.0, 80.0, 200.0))

    def test_06_line_1_plant_suffix_and_effective_date_isolation(self):
        self.assertEqual(normalize_line_code('Line 1'), '1')
        self.assertEqual(normalize_line_code('Line1'), '1')
        self.assertEqual(normalize_line_code('Line 1B'), '1B')
        self.assertNotEqual(normalize_line_code('Line 1'), normalize_line_code('Line 1B'))
        self.assertNotEqual(normalize_line_code('Line 1'), normalize_line_code('Line 13'))

        catalog = [{'factory': 'VH', 'plant': 'N14', 'line_raw': 'Line 1', 'source_group': 'Factory Sandals', 'effective_from': '2026-07-01', 'effective_to': '2026-08-31'}]
        rows = [
            # N14 Line 1 BEFORE effective date -> VH!
            {'date': '2026-06-15', 'factory': 'VH', 'plant': 'N14', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Cement', 'issue_quantity': 10},
            # N14 Line 1 IN effective date -> VB!
            {'date': '2026-07-15', 'factory': 'VH', 'plant': 'N14', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Cement', 'issue_quantity': 20},
            # N14 Line 1B (different suffix) -> MUST stay VH!
            {'date': '2026-07-15', 'factory': 'VH', 'plant': 'N14', 'line': 'Line 1B', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Cement', 'issue_quantity': 30},
            # S2 Line 1 (same line # different plant) -> MUST stay VH!
            {'date': '2026-07-15', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Cement', 'issue_quantity': 40},
            # JV Line 1 in P1 -> JV (or if catalog claimed JV, raises conflict!)
            {'date': '2026-07-15', 'factory': 'JV', 'plant': 'N14', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Cement', 'issue_quantity': 50},
        ]
        res, _, _ = _run_memory('FTT', rows, catalog_entries=catalog)
        g = res['top_groups'][0]
        self.assertEqual(g['vb_qty'], 20.0)
        self.assertEqual(g['vh_qty'], 80.0)
        self.assertEqual(g['jv_qty'], 50.0)
        self.assertEqual(g['clg_qty'], 150.0)

    def test_07_parent_subtype_detail_never_double_counted_and_control_mismatch_blocks(self):
        # Subtotals at row_level='parent' / 'subtype' must NOT be summed alongside row_level='detail'
        rows_ok = [
            {'audit_id': 'LOT1', 'row_level': 'parent', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'issue_quantity': 10},
            {'audit_id': 'LOT1', 'row_level': 'subtype', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Bond gap / Rat hole', 'issue_quantity': 10},
            {'audit_id': 'LOT1', 'row_level': 'detail', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Bond gap / Rat hole', 'detail_defect': 'Gap toe', 'issue_quantity': 6},
            {'audit_id': 'LOT1', 'row_level': 'detail', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'ftt_subtype': 'Bond gap / Rat hole', 'detail_defect': 'Gap heel', 'issue_quantity': 4},
        ]
        res, _, _ = _run_memory('FTT', rows_ok)
        self.assertEqual(res['top_groups'][0]['clg_qty'], 10.0)
        self.assertFalse(res['review']['has_blocking_issues'])

        # Detail > parent control in SAME unit -> must block!
        rows_bad = [
            {'audit_id': 'LOT2', 'row_level': 'parent', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'issue_quantity': 8, 'issue_unit': 'issues', 'control_unit': 'issues'},
            {'audit_id': 'LOT2', 'row_level': 'detail', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'detail_defect': 'Gap toe', 'issue_quantity': 10, 'issue_unit': 'issues', 'control_unit': 'issues'},
        ]
        res_bad, norm_bad, _ = _run_memory('FTT', rows_bad)
        self.assertTrue(norm_bad['has_blocking_mismatch'])
        self.assertTrue(res_bad['review']['has_blocking_issues'])
        self.assertEqual(res_bad['review']['status_code'], 'export_error')

    def test_08_multi_location_single_row_not_multiplied(self):
        rows = [
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Cement', 'detail_defect': 'Over cement', 'location': 'Toe; Quarter; Heel', 'issue_quantity': 10},
        ]
        res, _, _ = _run_memory('FTT', rows)
        self.assertEqual(res['top_groups'][0]['clg_qty'], 10.0)
        self.assertEqual(len(res['review']['multi_location_notes']), 1)

    def test_09_lot_repeated_across_multiple_details_never_duplicates_denominator(self):
        rows = [
            {'audit_id': 'AUD-001', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 200, 'defective_pairs': 12, 'standard_group': 'Cement', 'detail_defect': 'Over cement upper', 'issue_quantity': 8},
            {'audit_id': 'AUD-001', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 200, 'defective_pairs': 12, 'standard_group': 'Contamination', 'detail_defect': 'Ink stain', 'issue_quantity': 5},
            {'audit_id': 'AUD-001', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 200, 'defective_pairs': 12, 'standard_group': 'Bond gap / Bonding', 'detail_defect': 'Toe gap', 'issue_quantity': 3},
        ]
        res, _, _ = _run_memory('HFPA', rows)
        self.assertEqual(res['scope_totals']['VH']['record_count'], 1)
        self.assertEqual(res['scope_totals']['VH']['inspected_or_produced_qty'], 200.0)
        self.assertEqual(res['scope_totals']['VH']['defective_pairs_qty'], 12.0)

    def test_10_insufficient_detail_coverage_warns_and_keeps_full_group_denominator(self):
        rows = [
            {'audit_id': 'GAP1', 'row_level': 'parent', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 1000, 'standard_group': 'Bond gap / Bonding', 'issue_quantity': 1000},
            {'audit_id': 'GAP1', 'row_level': 'detail', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 1000, 'standard_group': 'Bond gap / Bonding', 'detail_defect': 'D1', 'issue_quantity': 400},
            {'audit_id': 'GAP1', 'row_level': 'detail', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 1000, 'standard_group': 'Bond gap / Bonding', 'detail_defect': 'D2', 'issue_quantity': 250},
            {'audit_id': 'GAP1', 'row_level': 'detail', 'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 1000, 'standard_group': 'Bond gap / Bonding', 'detail_defect': 'D3', 'issue_quantity': 150},
        ]
        res, _, _ = _run_memory('FTT', rows)
        cell = res['details_by_group']['bond_gap_bonding']['CLG']
        self.assertEqual(cell['group_total_qty'], 1000.0)
        self.assertEqual(cell['unreconciled_gap_qty'], 200.0)
        self.assertEqual(fmt_pct(cell['detail_coverage_pct']), '80.0%')
        self.assertEqual(cell['cell_status'], 'partial_coverage')
        self.assertIn('Top trong phần đã có chi tiết', cell['status_note'])
        # Denominator MUST be full 1,000, NOT sum(Top 3) = 800! (400/1000 = 40.0%, NOT 50.0%!)
        self.assertEqual(fmt_pct(cell['top_details'][0]['share_of_group_in_scope_pct']), '40.0%')

    def test_11_unmapped_label_goes_to_review_never_other(self):
        rows = [
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Brand New Strange Group', 'detail_defect': 'Mystery Flaw', 'issue_quantity': 75},
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Cement', 'detail_defect': 'Over cement', 'issue_quantity': 25},
        ]
        res, _, _ = _run_memory('FTT', rows)
        other_group = next(g for g in res['all_groups'] if g['group_id'] == 'other_miscellaneous')
        self.assertEqual(other_group['clg_qty'], 0.0)
        self.assertEqual(res['review']['status_code'], 'needs_mapping_review')
        self.assertEqual(len(res['review']['items']), 1)
        self.assertEqual(res['review']['items'][0]['raw_label'], 'Mystery Flaw')
        self.assertEqual(res['review']['items'][0]['affected_quantity'], 75.0)
        self.assertEqual(res['review']['items'][0]['status'], 'Chưa ánh xạ')

    def test_12_area_ranking_independent_from_clg_top3(self):
        rows = [
            # CLG: D1=500, D2=400, D3=300, D4=200 (so D4 is #4 in CLG)
            # In VH: D4=180, D3=50, D1=10, D2=5 -> VH Top 3 = [D4, D3, D1] (completely different from CLG [D1, D2, D3]!)
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M-VH', 'inspected_qty': 500, 'standard_group': 'Cement', 'detail_defect': 'D1', 'issue_quantity': 10},
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M-VH', 'inspected_qty': 500, 'standard_group': 'Cement', 'detail_defect': 'D2', 'issue_quantity': 5},
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M-VH', 'inspected_qty': 500, 'standard_group': 'Cement', 'detail_defect': 'D3', 'issue_quantity': 50},
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M-VH', 'inspected_qty': 500, 'standard_group': 'Cement', 'detail_defect': 'D4', 'issue_quantity': 180},
            {'date': '2026-06-05', 'factory': 'JV', 'plant': 'P1', 'line': 'Line 1', 'model': 'M-JV', 'inspected_qty': 500, 'standard_group': 'Cement', 'detail_defect': 'D1', 'issue_quantity': 490},
            {'date': '2026-06-05', 'factory': 'JV', 'plant': 'P1', 'line': 'Line 1', 'model': 'M-JV', 'inspected_qty': 500, 'standard_group': 'Cement', 'detail_defect': 'D2', 'issue_quantity': 395},
            {'date': '2026-06-05', 'factory': 'JV', 'plant': 'P1', 'line': 'Line 1', 'model': 'M-JV', 'inspected_qty': 500, 'standard_group': 'Cement', 'detail_defect': 'D3', 'issue_quantity': 250},
            {'date': '2026-06-05', 'factory': 'JV', 'plant': 'P1', 'line': 'Line 1', 'model': 'M-JV', 'inspected_qty': 500, 'standard_group': 'Cement', 'detail_defect': 'D4', 'issue_quantity': 20},
        ]
        res, _, _ = _run_memory('FTT', rows)
        clg_top = [d['detail_label'] for d in res['details_by_group']['cement']['CLG']['top_details']]
        vh_top = [d['detail_label'] for d in res['details_by_group']['cement']['VH']['top_details']]
        self.assertEqual(clg_top, ['D1', 'D2', 'D3'])
        self.assertEqual(vh_top, ['D4', 'D3', 'D1'])

    def test_13_hfpa_and_ftt_rankings_completely_independent(self):
        hfpa_rows = [{'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M_HFPA', 'inspected_qty': 100, 'standard_group': 'Contamination', 'issue_quantity': 90}]
        ftt_rows = [{'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M_FTT', 'inspected_qty': 100, 'standard_group': 'Cement', 'detail_defect': 'Over cement', 'issue_quantity': 80}]
        res_h, _, _ = _run_memory('HFPA', hfpa_rows)
        res_f, _, _ = _run_memory('FTT', ftt_rows)
        self.assertEqual(res_h['top_groups'][0]['group_id'], 'contamination')
        self.assertEqual(res_f['top_groups'][0]['group_id'], 'cement')

    def test_14_ties_zero_defects_missing_scope_and_na(self):
        rows = [
            # Exact tie at group rank 3: Airbag=10, Alignment=10, Contamination=10
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Cement', 'issue_quantity': 50},
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Bond gap / Bonding', 'issue_quantity': 30},
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_group_dummy': '', 'standard_group': 'Airbag', 'issue_quantity': 10},
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Contamination', 'issue_quantity': 10},
            # JV present with ZERO defects! (VB has NO source data -> missing_scope_data!)
            {'date': '2026-06-05', 'factory': 'JV', 'plant': 'P1', 'line': 'Line 1', 'model': 'M2', 'inspected_qty': 200, 'standard_group': 'Cement', 'issue_quantity': 0},
        ]
        res, _, _ = _run_memory('FTT', rows)
        self.assertTrue(res['tie_at_group_cutoff'])
        self.assertEqual(res['top_groups'][2]['group_id'], 'airbag')
        self.assertEqual(res['scope_totals']['JV']['scope_status'], 'no_defect')
        self.assertEqual(res['scope_totals']['VB']['scope_status'], 'missing_data')
        self.assertEqual(res['details_by_group']['cement']['JV']['cell_status'], 'no_defect')
        self.assertEqual(res['details_by_group']['cement']['VB']['cell_status'], 'missing_scope_data')
        self.assertEqual(fmt_pct(None), 'N/A')

    def test_15_internal_file_rejected_and_mapping_override_recalculates(self):
        with tempfile.TemporaryDirectory() as tmp:
            internal_path = Path(tmp) / 'FTT-Internal_2026-11-11.xlsx'
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'FTT-Internal'
            ws.append(['Date', 'FactoryGroup', 'Factory', 'Plant', 'Line', 'Model', 'Total Pair Produced', 'Metric', 'Defect Issues', "Issues Q'ty"])
            ws.append(['2026-06-05', 'CLG', 'VH', 'S2', 'Line 1', 'M1', 100, 'Rework1', 'Over cement', 5])
            wb.save(internal_path)
            wb.close()
            with self.assertRaises(ValueError):
                normalize_nike_source([internal_path], metric='FTT')

        # Override recalculation changes result and mapping_version!
        rows = [
            {'date': '2026-06-05', 'factory': 'VH', 'plant': 'S2', 'line': 'Line 1', 'model': 'M1', 'inspected_qty': 100, 'standard_group': 'Custom X', 'detail_defect': 'Custom Stain', 'issue_quantity': 40},
        ]
        m1 = load_defect_mapping()
        res1, _, _ = _run_memory('FTT', rows, mapping=m1)
        self.assertEqual(len(res1['top_groups']), 0)
        self.assertEqual(res1['review']['status_code'], 'needs_mapping_review')

        m2 = load_defect_mapping(
            extra_overrides=[
                {
                    'source_metric': 'FTT',
                    'raw_label': 'Custom Stain',
                    'standard_group_id': 'contamination',
                    'detail_label': 'Custom Stain',
                    'basis': 'Confirmed by user in test',
                }
            ]
        )
        self.assertNotEqual(m1['mapping_version'], m2['mapping_version'])
        res2, _, _ = _run_memory('FTT', rows, mapping=m2)
        self.assertEqual(res2['top_groups'][0]['group_id'], 'contamination')
        self.assertEqual(res2['top_groups'][0]['clg_qty'], 40.0)


if __name__ == '__main__':
    unittest.main()

