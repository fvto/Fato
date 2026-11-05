"""Run one imported report engine in its own process and private workspace."""
from pathlib import Path
import calendar
from dataclasses import asdict
import json
import os
import sys
import contextlib

from sfc_report_tools import file_hash, collect_files, make_bundle, now, read_json, write_json
from report_runtime import Monitor, Cancelled, diagnostics, friendly_error, QuietEngineOutput
from report_validation import validate_file, PreflightError
from report_exports import check_outputs
from report_brief import summarize_checks


def execute(folder):
    folder = Path(folder).resolve()
    run_file = folder / 'run.json'
    run = read_json(run_file)
    workspace = folder / 'workspace'
    output = workspace / 'Output'
    output.mkdir(exist_ok=True)
    os.chdir(workspace)
    sys.path.insert(0, str(workspace))
    sys.path.insert(0, str(workspace / 'Source'))
    run.update(status='running', pid=os.getpid(), stage='Đang kiểm tra đầu vào...')
    monitor = Monitor(folder, run)
    write_json(run_file, run)
    stage = monitor.stage

    try:
        tool, mode = run['tool'], run['mode']
        if run.get('acquire_missing'):
            stage('Thu thập đầu vào đúng tháng từ SFC', 'acquisition')
            if tool == 'quarterly_quality':
                from sfc_quarterly_quality_acquisition import acquire
            elif tool == 'reinspection':
                from sfc_reinspection_acquisition import acquire
            elif tool == 'hfpa':
                from sfc_hfpa_acquisition import acquire
            else:
                from sfc_copq_acquisition import acquire
            acquire(folder, run, stage)
            run['phase'] = 'processing'
        stage('Kiểm tra toàn bộ dữ liệu trước khi xử lý', 'validation')
        from sfc_report_tools import TOOLS, TOOLS_ROOT
        validation = []
        quarterly_sfc = tool == 'quarterly_quality' and run.get('acquire_missing')
        for index, item in enumerate(() if quarterly_sfc else run['inputs']):
            stage(f'Kiểm tra file {index+1}/{len(run["inputs"])}: {item["name"]}')
            path = workspace / TOOLS[tool]['groups'][item['group']][1] / item.get('staged_name', item['name'])
            original = next((r for r in run.get('validation', []) if r['group'] == item['group'] and
                            r.get('staged_name', r['name']) == path.name), None)
            if (tool == 'reinspection' and original and original['valid'] and
                    item.get('validated_sha256') and file_hash(path) == item['validated_sha256']):
                report = original
            else:
                report = validate_file(path, tool, item['group'], run['period'], normalize=True)
            validation.append(original if original and report['valid'] and tool != 'reinspection' else report)
            item['staged_name'] = report.get('staged_name', path.name)
            run['progress'] = dict(completed=index+1,total=len(run['inputs']),unit='file đã kiểm tra toàn bộ dòng')
            monitor.save()
            monitor.check_cancel()
        run['validation'] = validation
        if any(not report['valid'] for report in validation):
            raise PreflightError(validation)
        if tool == 'reinspection':
            import reinspection_suite as engine
            year, month = map(int, run['period'].split('-'))
            period_label = f'{calendar.month_name[month]} {year}'
            kwargs = dict(re_dir=str(workspace / 'Re-Inspection'), ftt_dir=str(workspace / 'FTT'),
                          hfpa_dir=str(workspace / 'HFPA'), month_label=period_label)
            stage('Đang tổng hợp dữ liệu Re-Inspection theo nhà máy...', 'metrics')
            metrics = engine.load_all_sites_data(
                kwargs['re_dir'], kwargs['ftt_dir'], kwargs['hfpa_dir'],
                log_fn=stage)
            kwargs['all_sites_data'] = metrics
            from sfc_report_tools import plant_of
            run['business_review'] = []
            for report in validation:
                if report['group'] != 'reinspection':
                    continue
                plant = plant_of(report['name'])
                actual = metrics.get(plant, {})
                expected = report.get('category_counts', {})
                for rule, expected_value, actual_value in [('rows',report['rows'],actual.get('total_records')),
                        ('qc',expected.get('QC',0),actual.get('qc_total_fail')),
                        ('ma',expected.get('MA',0),actual.get('ma_total_fail'))]:
                    run['business_review'].append(dict(rule='reinspection.'+plant+'.'+rule,
                        status='passed' if expected_value==actual_value else 'warning',
                        message='Đối chiếu dữ liệu Re-Inspection đã đọc · '+plant+' · '+rule,
                        expected=expected_value,actual=actual_value))
            if mode in ('all', 'database'):
                stage('Đang tạo Re-Inspection Database...', 'database')
                engine.generate_database_workbook(**kwargs,
                    output_file=str(output / f'Re-Inspection-Database-{run["period"]}.xlsx'),
                    color_template=str(workspace / 'Color_Template.xlsx'))
            if mode in ('all', 'recycle'):
                stage('Đang tạo Recycle Report...', 'recycle')
                engine.generate_recycle_workbook(**kwargs,
                    output_file=str(output / f'Recycle-Report-{run["period"]}.xlsx'))
            if mode in ('all', 'pptx'):
                stage('Đang tạo PowerPoint Re-Inspection...', 'presentation')
                engine.generate_presentation_report(**kwargs,
                    template_pptx=str(workspace / 'RE-INS REPORT.APR.2026.pptx'),
                    output_pptx=str(output / f'Re-Inspection-{run["period"]}.pptx'),
                    color_template=str(workspace / 'Color_Template.xlsx'))
        elif tool == 'copq':
            if mode == 'all':
                from agents.pipeline_orchestrator import PipelineOrchestratorAgent
                stage('Đang chạy COPQ → FTT B/C → Database → PowerPoint...')
                result = PipelineOrchestratorAgent(str(workspace)).execute_pipeline(period_override=run['period'], progress=stage)
                run['engine_warnings'] = len(result.get('warnings', []))
                if result.get('status') != 'SUCCESS':
                    raise RuntimeError('; '.join(result.get('errors', [])) or 'Pipeline COPQ chưa hoàn tất.')
            elif mode == 'combine':
                import combine_files
                stage('Đang gộp và phân tích COPQ Data...', 'copq_combine')
                # The isolated inputs already passed month validation by content.
                if not combine_files.combine_copq_files(input_dir=str(workspace / 'COPQ_Input'),
                        output_file=str(output / 'COPQ_Clean.xlsx')):
                    raise RuntimeError('Tool COPQ không tạo được workbook kết quả.')
            elif mode == 'ftt':
                import process_ftt
                stage('Đang tạo FTT B/C Report...', 'ftt')
                process_ftt.generate_ftt_report(input_folder=str(workspace / 'FTT_Input'),
                    output_file=str(output / 'FTT_Combined_Report.xlsx'),
                    color_map_file=str(workspace / 'BC_Color' / 'Color_Defect.xlsx'),
                    cost_file=str(workspace / 'Database' / 'CoPQ_type_analysis.xlsx'), cost_sheet_name='Current month')
        elif tool == 'hfpa':
            from agents.hfpa_agent_orchestrator import HFPAOrchestrator
            year, month = map(int, run['period'].split('-'))
            short_month = calendar.month_abbr[month]
            stage('Kiểm tra HFPA sẵn sàng', 'readiness')
            engine = HFPAOrchestrator(str(workspace))
            readiness = engine.ingestion_agent.evaluate_readiness(short_month, str(year))
            engine.ingestion_agent.print_readiness_report(readiness)
            (output / 'input_readiness.json').write_text(json.dumps(asdict(readiness), ensure_ascii=False, indent=2, default=str), encoding='utf8')
            if not readiness.is_ready:
                raise ValueError('Đầu vào HFPA chưa sẵn sàng; xem nhật ký và input_readiness.json.')
            if mode != 'check':
                stage('Đang đối chiếu QAStation và MES410...', 'reconciliation')
                from agents.color_manager import ColorManager
                ColorManager(str(workspace / 'Color_template.xlsx')).sync_reference_colors(
                    str(TOOLS_ROOT / 'reinspection' / 'Color_Template.xlsx')
                )
                inputs = {kind: [item.path for item in readiness.files if item.file_type == kind]
                          for kind in ('FTT', 'HFPA')}
                bus = engine.reconciliation_agent.run_reconciliation(short_month, str(year), input_files=inputs)
                stage('Đang cập nhật bản sao HFPA Database...', 'database')
                from populate_database import populate_hfpa_database
                if not populate_hfpa_database(bus):
                    raise RuntimeError('Chưa cập nhật được bản sao HFPA Database.')
                if mode == 'all':
                    stage('Đang tạo PowerPoint và bản phân tích chất lượng HFPA...', 'presentation')
                    if not engine.audit_agent.update_presentation_resiliently(bus):
                        raise RuntimeError('Chưa xuất được PowerPoint HFPA của lần chạy.')
                    report = engine.audit_agent.run_anomaly_audit(bus)
                    engine.audit_agent.generate_executive_brief(bus, report)
        else:
            # Additional local tools declare a Python adapter in registry metadata.
            import importlib
            adapter = TOOLS[tool].get('runner')
            if not adapter:
                raise ValueError('Tool chưa khai báo bộ xử lý.')
            if not quarterly_sfc:
                stage('Đang xử lý '+run['title'], 'processing')
            module, function = adapter.split(':',1)
            getattr(importlib.import_module(module), function)(workspace=workspace, run=run, stage=stage)
        stage('Đang kiểm tra và đóng gói kết quả...', 'export')
        files = collect_files(folder)
        for extension in TOOLS[tool].get('expected_extensions', {}).get(mode, []):
            if not any(item['name'].lower().endswith(extension) for item in files):
                raise RuntimeError('Thiếu định dạng kết quả bắt buộc: '+extension)
        run['export_review'] = check_outputs(folder, run, files)
        if run.get('engine_warnings'):
            run['export_review']['checks'].append(dict(rule='engine.warnings',status='warning',message='Bộ xử lý COPQ có cảnh báo ở bước database; cần kiểm tra trước khi sử dụng.'))
            run['export_review'].update(status='warning', warning_count=run['export_review']['warning_count']+1)
        run['export_review']['brief'] = summarize_checks(run['export_review']['checks'])
        run['requires_ack'] = run['export_review']['status'] == 'warning'
        if tool == 'hfpa' and mode != 'check':
            failures = [c['message'] for c in run['export_review']['checks']
                        if c['status'] != 'passed' and c['rule'] != 'hfpa.lot_variance']
            if failures:
                raise RuntimeError('Báo cáo HFPA chưa nhất quán. ' + ' '.join(failures))
        bundle = make_bundle(folder, files)
        monitor.check_cancel()
        history_status = ''
        if tool == 'copq' and mode in ('combine', 'all'):
            from copq_dashboard_history import update_history
            paths = [workspace / 'COPQ_Input' / item.get('staged_name', item['name'])
                     for item in run['inputs'] if item['group'] == 'copq']
            stage('Đang kiểm tra cập nhật file tổng COPQ...')
            history_status = update_history(paths, run['period'], verify_export_names=True)
        monitor.finish_step('done')
        run.update(status='done', stage='Đã hoàn tất; xem đối chiếu trước khi tải.' + (' ' + history_status if history_status else ''),
                   finished_at=now(), files=files, bundle=bundle)
    except Cancelled:
        monitor.finish_step('cancelled')
        run.update(status='cancelled', stage='Đã hủy tại điểm an toàn; giữ các file đã hoàn tất.', finished_at=now(),
                   files=collect_files(folder), requires_ack=True)
    except BaseException as error:
        monitor.finish_step('error')
        files = collect_files(folder)
        run.update(status='error', stage='Báo cáo chưa hoàn tất.', finished_at=now(), error=friendly_error(error),
                   diagnostics=diagnostics(error), files=files, requires_ack=True)
        if files:
            run['bundle'] = make_bundle(folder, files)
    write_json(run_file, run)
    return 0 if run['status'] == 'done' else 1


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf8', errors='replace')
    sys.stderr.reconfigure(encoding='utf8', errors='replace')
    with contextlib.redirect_stdout(QuietEngineOutput()), contextlib.redirect_stderr(QuietEngineOutput()):
        result = execute(sys.argv[1])
    raise SystemExit(result)
