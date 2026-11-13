import io
import json
import os
import tempfile
import time
import unittest
import zipfile
from pathlib import Path
from unittest import mock
from urllib.parse import unquote

import openpyxl

import report_tool_worker
import sfc_dashboard_server
import sfc_quarterly_defect_review_acquisition as qdr_acq
import sfc_report_tools
from dashboard_paths import APP_ROOT
from report_tools.quarterly_defect_review import report as qdr_report


HFPA_OFFICIAL_GROUPS = [
    'Airbag',
    'Alignment L+R Symmetry',
    'Bond gap / Bonding',
    'Cement',
    'Contamination',
    'Interior (including sock-liner)',
    'Poor Quality Spray / Paint',
    'Scratch / Tear / Rip / High Buffing',
    'Stitching / Loose Thread',
    'Wrinkle / Crease / Mis-shape',
    'Other / Miscellaneous',
]

FTT_TEST_SUBTYPES = [
    'Bond gap / Rat hole',
    'Delamination',
    'Wrinkle or Deformed Bottom',
    'Rocking',
    'Off-center',
    'Over cement',
    'Contamination',
]


def _make_hfpa_workbook(path: Path, records: list[dict], *, internal: bool = False) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'HFPA-Internal' if internal else 'HFPA-NIKE'
    if internal:
        ws.append(['Date', 'Factory', 'Plant', 'Line', 'Model', 'Sample', 'Defective Pairs', 'Report Mode'])
        ws.append(['2026-06-05', 'VH', 'S2', 'Line 1', 'M1', 100, 5, 'Internal'])
        wb.save(path)
        wb.close()
        return path

    headers = [
        'Factory Group Cd_UccId',
        'Audit Date (DD-MMM-YYYY)',
        'Internal Audit ID (YYMMDDno+Plant+Line)',
        'Plant',
        'Line',
        'Model',
        "Produced Q'ty",
        'Audit Sample Size',
        "Defective Q'ty(pair)",
        "Defect Q'ty (Total Defect Q'ty)",
        *HFPA_OFFICIAL_GROUPS,
    ]
    ws.append(headers)
    for rec in records:
        grp_counts = rec.get('groups', {})
        total_ctrl = sum(float(grp_counts.get(g, 0)) for g in HFPA_OFFICIAL_GROUPS)
        row = [
            rec.get('ucc_id', 'CL_VH'),
            rec['date'],
            rec['audit_id'],
            rec['plant'],
            rec['line'],
            rec['model'],
            rec.get('produced_qty', 400),
            rec['sample_qty'],
            rec['defective_pairs'],
            total_ctrl,
            *[float(grp_counts.get(g, 0)) for g in HFPA_OFFICIAL_GROUPS],
        ]
        ws.append(row)
    wb.save(path)
    wb.close()
    return path


def _make_ftt_workbook(path: Path, records: list[dict], *, internal: bool = False) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    if internal:
        ws.title = 'FTT-Internal'
        ws.append(['Date', 'FactoryGroup', 'Factory', 'Plant', 'Line', 'Model', 'Total Pair Produced', 'Metric', 'Defect Issues', "Issues Q'ty"])
        ws.append(['2026-06-05', 'CLG', 'VH', 'S2', 'Line 1', 'M1', 100, 'Rework1', 'Over cement', 5])
        wb.save(path)
        wb.close()
        return path

    ws.title = 'FTT Template'
    headers = [
        'Factory \nGroup ',
        'Factory Code',
        'Inspection Date (DD-MMM-YYYY)',
        'Plant',
        'Line',
        'Model',
        'Total A Grade Pairs',
        'Total B Grade Pairs (Reject)',
        'Total C Grade Pairs (Reject)',
        'Total Rework',
        '1st Time Rework Pairs',
        *[f'{sub}_Rework' for sub in FTT_TEST_SUBTYPES],
    ]
    ws.append(headers)
    for rec in records:
        sub_counts = rec.get('subtypes', {})
        total_rw = sum(float(sub_counts.get(s, 0)) for s in FTT_TEST_SUBTYPES)
        row = [
            rec.get('factory_group', 'CLG'),
            rec.get('factory', 'VH'),
            rec['date'],
            rec['plant'],
            rec['line'],
            rec['model'],
            rec.get('a_grade', 1000),
            rec.get('b_grade', 2),
            rec.get('c_grade', 1),
            total_rw,
            total_rw,
            *[float(sub_counts.get(s, 0)) for s in FTT_TEST_SUBTYPES],
        ]
        ws.append(row)
    wb.save(path)
    wb.close()
    return path


class QuarterlyDefectReviewIntegrationTests(unittest.TestCase):
    def test_01_sfc_acquisition_plan_html_parser_and_isolation(self):
        qa_html = """
        <html><body>
          <table>
            <tr><td><a href="QADetail.aspx?factory=VT&mode=nike">Factory Sandals (VB)</a></td></tr>
            <tr><td><a href="QADetail.aspx?factory=S2">Factory S2</a></td></tr>
          </table>
        </body></html>
        """
        href = qdr_acq.parse_factory_sandals_href(qa_html)
        self.assertEqual(href, 'QADetail.aspx?factory=VT&mode=nike')

        detail_html = """
        <html><body>
          <table id="gvDetail">
            <tr><th>Plant</th><th>Line</th><th>Link</th></tr>
            <tr><td>Total</td><td>Total</td><td>Total</td></tr>
            <tr><td>N14</td><td>Line 1</td><td><a href="FTT.aspx?line=N14%20Line%201">N14 Line 1</a></td></tr>
            <tr><td>VB</td><td>PCC</td><td><a href="FTT.aspx?line=VB%20PCC">VB PCC</a></td></tr>
            <tr><td>Factory</td><td>Sandals</td><td>Factory Sandals</td></tr>
          </table>
        </body></html>
        """
        parsed = qdr_acq.parse_factory_sandals_html(detail_html)
        pairs = {(e['plant'], e['line_std']) for e in parsed}
        self.assertEqual(pairs, {('N14', '1'), ('VB', 'PCC')})
        self.assertTrue(all(e['report_area'] == 'VB' and e['source_group'] == 'Factory Sandals' for e in parsed))

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            cache_root = tmp_path / 'IsolatedCache'
            acq_plan = qdr_acq.plan(period='FY27-Q1', date_from='2026-06-01', date_to='2026-06-30', cache_root=cache_root)
            self.assertEqual(acq_plan['workflow'], 'isolated_nike_review')
            self.assertEqual(acq_plan['available'], [])
            self.assertEqual(len(acq_plan['missing']), 8)  # 4 plants x (hfpa_nike + ftt_nike)

            # Seed mock SFC acquisition & confirm strict NIKE flags + partial resilience
            ws_dir = tmp_path / 'Workspace'
            ws_dir.mkdir(parents=True, exist_ok=True)
            seen_hfpa_flags = []
            seen_ftt_payloads = []
            seen_navigations = []

            class FakeSession:
                def __init__(self, plant, cookies=None):
                    self.plant = plant

                def prepare(self):
                    return self

                def request(self, url, params=None):
                    if 'factory=VT' in url:
                        return detail_html.encode('utf-8')
                    return qa_html.encode('utf-8')

                def export_period(self, d_from, d_to, nike, stem):
                    seen_hfpa_flags.append((self.plant, nike))
                    if self.plant == 'VH2':
                        raise RuntimeError('Simulated VH2 network glitch')
                    out = Path(f'{stem}.xlsx')
                    out.parent.mkdir(parents=True, exist_ok=True)
                    _make_hfpa_workbook(
                        out,
                        [{
                            'date': d_from,
                            'audit_id': f'260601{self.plant}',
                            'ucc_id': f'CL_{self.plant}',
                            'plant': 'S2' if self.plant == 'VH' else 'P1',
                            'line': 'Line 1',
                            'model': 'NIKE WINFLO 12',
                            'sample_qty': 100,
                            'defective_pairs': 5,
                            'groups': {'Bond gap / Bonding': 5},
                        }],
                    )
                    yield {'file': {'path': str(out)}}

            class FakeController:
                def switch_factory(self, plant):
                    self.plant = plant
                    return {'success': True, 'current_plant': plant}

                def apply_sfc_filters(self, **params):
                    self.asserted_nike_mode = params['nike']
                    session = FakeSession(self.plant)
                    events = list(session.export_period(
                        params['date_from'], params['date_to'], params['nike'], params['export_path'],
                    ))
                    return {'success': True, 'file': events[-1]['file']}

                def navigate_to(self, url):
                    seen_navigations.append(url)
                    return {'success': True, 'url': url}

                def set_selected_factories(self, plants):
                    return None

                def execute_dynamic_action(self, action, target, payload):
                    if action == 'ftt_station_export':
                        seen_ftt_payloads.append(dict(payload))
                        plant = payload['_expected_plant']
                        out = tmp_path / f'dl_{plant}_{payload["date_from"]}.xlsx'
                        _make_ftt_workbook(
                            out,
                            [{
                                'factory_group': 'CLG',
                                'factory': plant,
                                'date': payload['date_from'],
                                'plant': 'S2' if plant.startswith('VH') else 'P1',
                                'line': 'Line 1',
                                'model': 'NIKE WINFLO 12',
                                'subtypes': {'Bond gap / Rat hole': 4},
                            }],
                        )
                        return {'success': True, 'downloaded_file': {'path': str(out)}}
                    return {'success': True}

            run_obj = {'period': 'FY27-Q1', 'date_from': '2026-06-01', 'date_to': '2026-06-30'}
            acq_res = qdr_acq.acquire(
                ws_dir,
                run_obj,
                lambda _m, _p=None: None,
                controller=FakeController(),
                session_factory=FakeSession,
                cache_root=cache_root,
            )

            self.assertTrue(all(flag is True for _p, flag in seen_hfpa_flags))
            self.assertTrue(seen_navigations)
            self.assertTrue(all('QA.aspx?tree_id=Q602' in u for u in seen_navigations))
            self.assertTrue(all(p['report_type'] == 'nike' for p in seen_ftt_payloads))
            self.assertTrue(all(p.get('_expected_plant') in {'VH', 'VH2', 'JV', 'JV2'} for p in seen_ftt_payloads))
            self.assertTrue(all(str(p.get('file_name', '')).endswith('_2026-06-30') for p in seen_ftt_payloads))
            self.assertEqual(len(acq_res['errors']), 1)  # VH2 partial error preserved cleanly
            self.assertEqual(len(acq_res['downloaded']), 7)  # 3 HFPA + 4 FTT downloaded cleanly
            self.assertGreaterEqual(len(list((ws_dir / 'Inputs' / 'hfpa_nike').glob('*.xlsx'))), 3)
            self.assertEqual(len(list((ws_dir / 'Inputs' / 'ftt_nike').glob('*.xlsx'))), 4)
            self.assertTrue((ws_dir / 'Inputs' / 'sandals_catalog' / 'line_catalog_factory_sandals.json').is_file())
            self.assertEqual(len(list((tmp_path).glob('SharedExport*'))), 0)

            # Confirm sfc_excel.file_info works for isolated cache paths outside EXPORT_ROOT
            import sfc_excel
            iso_sample = cache_root / 'HFPA_NIKE' / 'VH' / 'sample.xlsx'
            iso_sample.parent.mkdir(parents=True, exist_ok=True)
            _make_hfpa_workbook(
                iso_sample,
                [{
                    'date': '2026-06-05',
                    'audit_id': '260605VH1',
                    'plant': 'S2',
                    'line': 'Line 1',
                    'model': 'NIKE WINFLO 12',
                    'sample_qty': 50,
                    'defective_pairs': 2,
                    'groups': {'Bond gap / Bonding': 2},
                }],
            )
            info = sfc_excel.file_info(iso_sample, rows=1)
            self.assertEqual(info['url'], '')
            self.assertEqual(info['rows'], 1)

    def test_02_real_excel_pivot_formulas_and_wps_reopen_e2e(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            # Construct multi-month FY27-Q1 dataset across VH, JV, VB + unmatched CL_VB test case
            hfpa_rows = [
                # June
                {'ucc_id': 'CL_VH', 'date': '2026-06-10', 'audit_id': '260610VH1', 'plant': 'S2', 'line': 'Line 1', 'model': 'NIKE WINFLO 12', 'sample_qty': 1000, 'defective_pairs': 50, 'groups': {'Bond gap / Bonding': 60, 'Contamination': 20}},
                {'ucc_id': 'CL_VB', 'date': '2026-06-11', 'audit_id': '260611VB1', 'plant': 'N14', 'line': 'Line 1', 'model': 'JORDAN FRANCHISE SLIDE', 'sample_qty': 500, 'defective_pairs': 30, 'groups': {'Bond gap / Bonding': 40, 'Cement': 10}},
                {'ucc_id': 'CL_VB', 'date': '2026-06-12', 'audit_id': '260612UN1', 'plant': 'S9', 'line': 'Line 9', 'model': 'UNMATCHED VB MODEL', 'sample_qty': 200, 'defective_pairs': 10, 'groups': {'Contamination': 10}},
                # July
                {'ucc_id': 'CL_JV', 'date': '2026-07-15', 'audit_id': '260715JV1', 'plant': 'P1', 'line': 'Line 2', 'model': 'JORDAN SPIZIKE LOW (GS)', 'sample_qty': 800, 'defective_pairs': 40, 'groups': {'Bond gap / Bonding': 50, 'Wrinkle / Crease / Mis-shape': 30}},
                # August
                {'ucc_id': 'CL_VH', 'date': '2026-08-20', 'audit_id': '260820VH2', 'plant': 'S2', 'line': 'Line 2', 'model': 'JORDAN SPIZIKE LOW (PS)', 'sample_qty': 600, 'defective_pairs': 25, 'groups': {'Contamination': 35, 'Cement': 15}},
            ]
            ftt_rows = [
                # VH (S2 Line 1 & Line 2)
                {'factory_group': 'CLG', 'factory': 'VH', 'date': '2026-06-10', 'plant': 'S2', 'line': 'Line 1', 'model': 'NIKE WINFLO 12', 'subtypes': {'Bond gap / Rat hole': 120, 'Delamination': 30, 'Wrinkle or Deformed Bottom': 40, 'Over cement': 15}},
                {'factory_group': 'CLG', 'factory': 'VH', 'date': '2026-08-20', 'plant': 'S2', 'line': 'Line 2', 'model': 'JORDAN SPIZIKE LOW (GS)', 'subtypes': {'Rocking': 60, 'Contamination': 25}},
                # VB (N14 Line 1 tagged Factory Sandals)
                {'factory_group': 'Factory Sandals', 'factory': 'VH', 'date': '2026-06-11', 'plant': 'N14', 'line': 'Line 1', 'model': 'JORDAN FRANCHISE SLIDE', 'subtypes': {'Delamination': 80, 'Bond gap / Rat hole': 20, 'Off-center': 35}},
                # JV (P1 Line 2)
                {'factory_group': 'CLG', 'factory': 'JV', 'date': '2026-07-15', 'plant': 'P1', 'line': 'Line 2', 'model': 'JORDAN SPIZIKE LOW (PS)', 'subtypes': {'Bond gap / Rat hole': 100, 'Delamination': 50, 'Wrinkle or Deformed Bottom': 45, 'Over cement': 30}},
            ]

            hfpa_file = _make_hfpa_workbook(tmp_path / 'HFPA_NIKE_Q1.xlsx', hfpa_rows)
            ftt_file = _make_ftt_workbook(tmp_path / 'FTT-NIKE_Q1.xlsx', ftt_rows)
            out_dir = tmp_path / 'Output'

            summary = qdr_report.run_quarterly_defect_review(
                hfpa_inputs=[hfpa_file],
                ftt_inputs=[ftt_file],
                output_dir=out_dir,
                period='FY27-Q1',
                date_from='2026-06-01',
                date_to='2026-08-31',
                run_id='integration_e2e',
                skip_office=False,
                fallback_catalog_path=None,
            )

            # 1. Independent Python recomputation verification (Section 14)
            # Bonding FTT: VH=(120+30)=150, VB=(80+20)=100, JV=(100+50)=150 -> CLG=400
            ftt_top1 = summary['ftt']['top_groups'][0]
            self.assertEqual(ftt_top1['group_id'], 'bond_gap_bonding')
            self.assertAlmostEqual(ftt_top1['vh_qty'], 150.0)
            self.assertAlmostEqual(ftt_top1['jv_qty'], 150.0)
            self.assertAlmostEqual(ftt_top1['vb_qty'], 100.0)
            self.assertAlmostEqual(ftt_top1['clg_qty'], 400.0)
            self.assertAlmostEqual(ftt_top1['vh_contrib_pct'], 150.0 * 100 / 400.0, places=3)

            # HFPA: unmatched S9 Line 9 UccId=CL_VB stays in VH (10 Contamination) with non-silent review warning!
            hfpa_vh_total = summary['hfpa']['scope_totals']['VH']['total_defect_issues']
            hfpa_vb_total = summary['hfpa']['scope_totals']['VB']['total_defect_issues']
            self.assertAlmostEqual(hfpa_vh_total, 60 + 20 + 10 + 35 + 15)  # 140
            self.assertAlmostEqual(hfpa_vb_total, 40 + 10)  # 50
            self.assertTrue(any(it['review_type'] == 'hfpa_uccid_catalog_mismatch' for it in summary['hfpa']['review']['items']))

            # FTT Detail denominator check: VB Bonding top detail is Delamination = 80/100 (80.0% of VB group), area contrib = 80/160 (50.0% of CLG Delamination)
            vb_bonding_top = summary['ftt']['details_by_group']['bond_gap_bonding']['VB']['top_details'][0]
            self.assertEqual(vb_bonding_top['detail_label'], 'Delamination')
            self.assertAlmostEqual(vb_bonding_top['share_of_group_in_scope_pct'], 80.0, places=3)
            self.assertAlmostEqual(vb_bonding_top['area_contribution_to_detail_pct'], 50.0, places=3)

            # 2. Excel workbook structure & dynamic formula check (5 required sheets + sorted M:N + Slide 1 dynamic formulas)
            excel_path = Path(summary['files']['excel'])
            self.assertTrue(excel_path.is_file())
            wb = openpyxl.load_workbook(excel_path, data_only=False)
            try:
                self.assertEqual(
                    wb.sheetnames,
                    [
                        'Overview_and_Charts',
                        'Drilldown_Normalized',
                        'Detail_Breakdown_HFPA',
                        'Detail_Breakdown_FTT',
                        'Catalog_Mapping_&_Review',
                    ],
                )
                ws_over = wb['Overview_and_Charts']
                formulas_over = [
                    str(c.value)
                    for row in ws_over.iter_rows(min_row=1, max_row=45, min_col=1, max_col=18)
                    for c in row
                    if isinstance(c.value, str) and c.value.startswith('=')
                ]
                self.assertTrue(any('SUMIFS(' in f for f in formulas_over))
                self.assertTrue(any('AVERAGE(' in f for f in formulas_over))
                self.assertTrue(any('SUM(' in f for f in formulas_over))

                ws_ftt = wb['Detail_Breakdown_FTT']
                chart_vals = [
                    float(ws_ftt.cell(r, 14).value)
                    for r in range(4, 25)
                    if ws_ftt.cell(r, 13).value and isinstance(ws_ftt.cell(r, 14).value, (int, float))
                ]
                self.assertGreaterEqual(len(chart_vals), 2)
                self.assertEqual(chart_vals, sorted(chart_vals, reverse=True))
            finally:
                wb.close()

            # 3. Office/WPS COM native export & reopen check (AGENTS.md Case 17)
            self.assertTrue(summary['export_ok'], summary.get('export_error'))
            office_meta = summary['office_meta']
            self.assertTrue(office_meta['reopen_verified'])
            self.assertIn(office_meta['verification_engine'], ('wps', 'msoffice'))
            self.assertEqual(office_meta['slide_count'], 7)
            self.assertGreaterEqual(office_meta['native_charts_pasted'], 5)

            pptx_path = Path(summary['files']['pptx'])
            self.assertTrue(pptx_path.is_file())
            ext_links = office_meta.get('external_links', [])
            self.assertGreaterEqual(len(ext_links), 1)
            for link in ext_links:
                target = str(link['target'] if isinstance(link, dict) else link)
                if target.startswith('file:///'):
                    self.assertEqual(Path(unquote(target[8:])).resolve(), excel_path.resolve())
                else:
                    self.assertIn(excel_path.name.lower(), target.lower())

            with zipfile.ZipFile(pptx_path, 'r') as zf:
                names = zf.namelist()
                slide_xmls = [n for n in names if n.startswith('ppt/slides/slide') and n.endswith('.xml')]
                self.assertEqual(len(slide_xmls), 7)

            previews = summary['files']['slide_previews']
            self.assertEqual(len(previews), 7)
            for p_str in previews:
                p = Path(p_str)
                self.assertTrue(p.is_file() and p.stat().st_size > 1000)

    def test_03_flask_api_validation_run_and_isolated_overrides(self):
        tableau_path = APP_ROOT / 'Report' / 'Quarterly' / 'Tableau data source 2026 HFPA Bonding.xlsx'
        tableau_stat_before = tableau_path.stat().st_mtime_ns if tableau_path.exists() else None

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            tmp_runs = tmp_path / 'ToolRuns'
            tmp_cache = tmp_path / 'CacheQDR'
            tmp_overrides = tmp_cache / 'mapping_overrides.json'
            tmp_cache.mkdir(parents=True, exist_ok=True)

            with (
                mock.patch.object(sfc_report_tools, 'RUN_ROOT', tmp_runs),
                mock.patch.object(qdr_acq, 'ISOLATED_CACHE_ROOT', tmp_cache),
                mock.patch.object(qdr_report, 'DEFAULT_ISOLATED_CACHE_DIR', tmp_cache),
                mock.patch.object(qdr_report, 'DEFAULT_CATALOG_CACHE_PATH', tmp_cache / 'line_catalog_factory_sandals.json'),
                mock.patch.object(qdr_report, 'DEFAULT_MAPPING_OVERRIDES_PATH', tmp_overrides),
            ):
                client = sfc_dashboard_server.app.test_client()

                # 1. Registry defaults & automatic-mode preflight validation check (prevents cutoff duplicate kwarg!)
                r_auto_val = client.post(
                    '/api/report-tools/quarterly_defect_review/validate',
                    data={
                        'period': 'FY27-Q1',
                        'mode': 'all',
                        'date_from': '2026-06-01',
                        'date_to': '2026-08-31',
                        'cutoff': '',
                        'selections': '[]',
                        'acquire_missing': 'true',
                    },
                    content_type='multipart/form-data',
                )
                self.assertEqual(r_auto_val.status_code, 200, r_auto_val.get_data(as_text=True))
                auto_val_json = r_auto_val.get_json()
                self.assertTrue(auto_val_json['valid'])
                self.assertGreater(len(auto_val_json['will_download']), 0)

                # 1. Registry defaults check
                r_tools = client.get('/api/report-tools')
                self.assertEqual(r_tools.status_code, 200)
                tools_map = {t['id']: t for t in r_tools.get_json()['tools']}
                qdr_tool = tools_map['quarterly_defect_review']
                self.assertEqual(qdr_tool['default_period'], 'FY27-Q1')
                self.assertEqual(qdr_tool['default_date_from'], '2026-06-01')
                self.assertEqual(qdr_tool['default_date_to'], '2026-08-31')

                # 2. Reject Internal file cleanly (Case 16)
                internal_buf = io.BytesIO()
                p_int = _make_ftt_workbook(tmp_path / 'FTT-Internal.xlsx', [], internal=True)
                internal_buf.write(p_int.read_bytes())
                internal_buf.seek(0)
                r_bad = client.post(
                    '/api/report-tools/quarterly_defect_review/validate',
                    data={
                        'period': 'FY27-Q1',
                        'mode': 'excel',
                        'date_from': '2026-06-01',
                        'date_to': '2026-08-31',
                        'acquire_missing': 'false',
                        'ftt_nike': (internal_buf, 'FTT-Internal.xlsx'),
                    },
                    content_type='multipart/form-data',
                )
                self.assertEqual(r_bad.status_code, 422)
                self.assertFalse(r_bad.get_json()['validation'][0]['valid'])

                # 3. Validate valid NIKE files & execute background run (mode='excel')
                p_hfpa = _make_hfpa_workbook(
                    tmp_path / 'HFPA_VH_NIKE.xlsx',
                    [{
                        'ucc_id': 'CL_VH',
                        'date': '2026-06-05',
                        'audit_id': '260605VH1',
                        'plant': 'S2',
                        'line': 'Line 1',
                        'model': 'NIKE WINFLO 12',
                        'sample_qty': 100,
                        'defective_pairs': 5,
                        'groups': {'Bond gap / Bonding': 6},
                    }],
                )
                p_ftt = _make_ftt_workbook(
                    tmp_path / 'FTT-NIKE.xlsx',
                    [{
                        'factory_group': 'CLG',
                        'factory': 'VH',
                        'date': '2026-06-05',
                        'plant': 'S2',
                        'line': 'Line 1',
                        'model': 'NIKE WINFLO 12',
                        'subtypes': {'Bond gap / Rat hole': 4},
                    }],
                )
                def _fake_popen(cmd, *args, **kwargs):
                    run_folder_arg = Path(cmd[-1])
                    prev_cwd = os.getcwd()
                    try:
                        report_tool_worker.execute(run_folder_arg)
                    finally:
                        os.chdir(prev_cwd)

                    class _DoneProc:
                        pid = 999999
                        returncode = 0

                        def poll(self):
                            return 0

                    return _DoneProc()

                with mock.patch.object(sfc_report_tools.subprocess, 'Popen', side_effect=_fake_popen):
                    r_run = client.post(
                        '/api/report-tools/quarterly_defect_review/runs',
                        data={
                            'period': 'FY27-Q1',
                            'mode': 'excel',
                            'date_from': '2026-06-01',
                            'date_to': '2026-08-31',
                            'acquire_missing': 'false',
                            'hfpa_nike': (io.BytesIO(p_hfpa.read_bytes()), 'HFPA_VH_NIKE.xlsx'),
                            'ftt_nike': (io.BytesIO(p_ftt.read_bytes()), 'FTT-NIKE.xlsx'),
                        },
                        content_type='multipart/form-data',
                    )
                self.assertIn(r_run.status_code, (200, 201, 202))
                run_id = r_run.get_json()['id']

                polled = client.get(f'/api/report-tools/runs/{run_id}').get_json()
                self.assertEqual(polled.get('status'), 'done', polled)
                r_sum = client.get(f'/api/report-tools/runs/{run_id}/defect-review')
                self.assertEqual(r_sum.status_code, 200)
                rev_sum = r_sum.get_json()['review_summary']
                self.assertEqual(rev_sum['hfpa']['scope_totals']['VH']['total_defect_issues'], 6.0)
                self.assertEqual(rev_sum['ftt']['scope_totals']['VH']['total_defect_issues'], 4.0)

                # 4. Isolated mapping overrides endpoint check
                r_ov = client.post(
                    '/api/report-tools/quarterly_defect_review/overrides',
                    json={
                        'rules': [{
                            'source_metric': 'FTT',
                            'raw_label': 'Custom Integration Flap',
                            'standard_group_id': 'bond_gap_bonding',
                            'confirmed_by': 'integration_test',
                        }]
                    },
                )
                self.assertEqual(r_ov.status_code, 200)
                self.assertTrue(tmp_overrides.is_file())
                r_get_ov = client.get('/api/report-tools/quarterly_defect_review/overrides')
                self.assertEqual(r_get_ov.status_code, 200)
                self.assertEqual(len(r_get_ov.get_json()['rules']), 1)

        if tableau_stat_before is not None:
            self.assertEqual(tableau_path.stat().st_mtime_ns, tableau_stat_before)


if __name__ == '__main__':
    unittest.main()
