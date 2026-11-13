"""Native Office / WPS Copy/Paste PPTX assembly and reopen verification (AGENTS.md compliant)."""

from __future__ import annotations

import shutil
import tempfile
import time
from pathlib import Path
from typing import Any
from urllib.parse import unquote
from zipfile import ZipFile

from lxml import etree
from PIL import ImageGrab


class OfficeExportError(RuntimeError):
    """Raised when Excel/WPS COM chart transfer or PPTX reopen verification fails."""


_DISCONNECTED = {-2147023174, -2147417848}  # RPC server unavailable / object disconnected
_TRANSIENT_COM = _DISCONNECTED | {-2147418111, -2147417846}


def _com_code(error: Exception) -> int | None:
    code = getattr(error, 'hresult', None)
    if code is None and error.args and isinstance(error.args[0], int):
        code = error.args[0]
    return code


def _quit_after_close(app: Any) -> None:
    """WPS may already have ended this proxy when the sibling app quit."""
    try:
        app.Quit()
    except AttributeError:
        # WPS can remove the late-bound Quit member after its sibling has
        # already exited. All owned documents were saved/closed beforehand;
        # the fresh reopen below still has to pass verification.
        return
    except Exception as error:
        if _com_code(error) not in _DISCONNECTED:
            raise


def _close_saved_document(document: Any, *args: Any) -> None:
    """A saved WPS document may already be closed by the shared app lifecycle."""
    try:
        document.Close(*args)
    except Exception as error:
        if _com_code(error) not in _DISCONNECTED:
            raise


def _open_verification_session(client: Any, path: Path) -> tuple[Any, Any]:
    """Acquire a fresh COM proxy after shutting down Excel/WPS, never reuse a dead one."""
    for attempt in range(3):
        app = None
        try:
            app = client.DispatchEx('PowerPoint.Application')
            app.DisplayAlerts = 0
            return app, app.Presentations.Open(str(path), True, False, False)
        except Exception as error:
            if app is not None:
                try:
                    _quit_after_close(app)
                except Exception:
                    pass
            if _com_code(error) not in _TRANSIENT_COM or attempt == 2:
                raise
            time.sleep(0.5)


def _retry(action: Any, attempts: int = 8, delay: float = 0.25) -> Any:
    last_error: Exception | None = None
    for idx in range(attempts):
        try:
            return action()
        except Exception as exc:
            last_error = exc
            if idx < attempts - 1:
                time.sleep(delay)
    raise last_error  # type: ignore[misc]


def _detect_app_label(excel_app: Any, ppt_app: Any) -> dict[str, str]:
    xl_path = str(getattr(excel_app, 'Path', '') or '')
    ppt_path = str(getattr(ppt_app, 'Path', '') or '')
    combined = f'{xl_path} {ppt_path}'.lower()
    if 'kingsoft' in combined or 'wps' in combined:
        label = 'WPS Office (WPS Spreadsheets & WPS Presentation)'
        engine = 'wps'
    else:
        label = 'Microsoft Office (Microsoft Excel & Microsoft PowerPoint)'
        engine = 'msoffice'
    return {'engine': engine, 'label': label, 'excel_path': xl_path, 'ppt_path': ppt_path}


def _copy_range_image(sheet: Any, rng_ref: str, png_path: Path) -> tuple[int, int]:
    """Capture an exact bitmap snapshot from Excel via CopyPicture without post-editing."""
    sheet.Activate()

    def _grab() -> Any:
        sheet.Range(rng_ref).CopyPicture(1, 2)  # xlScreen=1, xlBitmap=2
        time.sleep(0.2)
        img = ImageGrab.grabclipboard()
        if img is None:
            raise OfficeExportError(f'Không lấy được ảnh bảng Excel {sheet.Name}!{rng_ref} từ clipboard.')
        return img

    img = _retry(_grab, attempts=6, delay=0.3)
    img.save(png_path)
    return img.width, img.height


def _add_fitted_picture(slide: Any, png_path: Path, box: tuple[float, float, float, float], w_px: int, h_px: int) -> Any:
    left, top, max_w, max_h = box
    ratio = min(max_w / max(1, w_px), max_h / max(1, h_px))
    draw_w = w_px * ratio
    draw_h = h_px * ratio
    return slide.Shapes.AddPicture(str(png_path), False, True, left, top, draw_w, draw_h)


def _paste_chart(sheet: Any, chart_obj: Any, slide: Any, box: tuple[float, float, float, float], name: str) -> Any:
    """Copy a finalized Excel ChartObject natively and adjust ONLY Left/Top/Width/Height/Name in PPT."""
    def _do() -> Any:
        sheet.Activate()
        chart_obj.Copy()
        time.sleep(0.25)
        try:
            pasted = slide.Shapes.PasteSpecial(DataType=0, Link=False)
        except Exception:
            pasted = slide.Shapes.Paste()
        shape = pasted.Item(1) if hasattr(pasted, 'Count') else pasted
        if not getattr(shape, 'HasChart', False):
            raise OfficeExportError(f'Đối tượng dán {name} không phải native Chart.')
        return shape

    shape = _retry(_do, attempts=6, delay=0.35)
    shape.Left, shape.Top, shape.Width, shape.Height = box
    try:
        shape.Name = name
    except Exception:
        pass
    return shape


def _add_slide_header(slide: Any, title: str, subtitle: str, slide_w: float) -> None:
    hdr = slide.Shapes.AddShape(1, 0, 0, slide_w, 46)
    hdr.Fill.ForeColor.RGB = 0x5D361B  # BGR for #1B365D! (COM RGB is 0xBBGGRR)
    hdr.Line.Visible = 0
    tf = hdr.TextFrame.TextRange
    tf.Text = title
    tf.Font.Name = 'Calibri'
    tf.Font.Size = 18
    tf.Font.Bold = -1
    tf.Font.Color.RGB = 0xFFFFFF

    if subtitle:
        sub = slide.Shapes.AddTextbox(1, 24, 50, slide_w - 48, 24)
        stf = sub.TextFrame.TextRange
        stf.Text = subtitle
        stf.Font.Name = 'Calibri'
        stf.Font.Size = 10
        stf.Font.Color.RGB = 0x695547  # BGR for #475569


def _add_note_box(slide: Any, box: tuple[float, float, float, float], lines: list[str], *, fill_bgr: int = 0xFEFAF8) -> None:
    left, top, width, height = box
    shape = slide.Shapes.AddShape(1, left, top, width, height)
    shape.Fill.ForeColor.RGB = fill_bgr
    shape.Line.ForeColor.RGB = 0xE1D5CB
    shape.TextFrame.WordWrap = -1
    shape.TextFrame.MarginLeft = 14
    shape.TextFrame.MarginRight = 14
    shape.TextFrame.MarginTop = 8
    shape.TextFrame.MarginBottom = 8
    tr = shape.TextFrame.TextRange
    tr.Text = '\n'.join(lines)
    tr.Font.Name = 'Calibri'
    tr.Font.Size = 10
    tr.Font.Color.RGB = 0x3B291E
    try:
        tr.ParagraphFormat.Alignment = 1  # ppAlignLeft
    except Exception:
        pass


def _inspect_pptx_package(pptx_path: Path, final_xlsx_path: Path) -> dict[str, Any]:
    """Read-only inspection of PPTX zip ensuring zero XML patching and no dead temp links."""
    resolved_xlsx = final_xlsx_path.resolve()
    chart_parts: list[str] = []
    external_links: list[dict[str, str]] = []
    dead_or_temp_links: list[str] = []

    with ZipFile(pptx_path) as zf:
        for name in zf.namelist():
            if name.startswith('ppt/charts/chart') and name.endswith('.xml'):
                chart_parts.append(name)
            if name.endswith('.rels'):
                root = etree.fromstring(zf.read(name))
                for rel in root:
                    if rel.get('TargetMode') == 'External':
                        target = rel.get('Target', '')
                        external_links.append({'rels_part': name, 'target': target})
                        if target.startswith('file:///'):
                            local = unquote(target[8:])
                            try:
                                resolved_target = Path(local).resolve()
                            except OSError:
                                dead_or_temp_links.append(target)
                            else:
                                if resolved_target != resolved_xlsx:
                                    dead_or_temp_links.append(target)
                        else:
                            low = target.lower()
                            if 'temp' in low or 'tmp' in low or resolved_xlsx.name.lower() not in low:
                                dead_or_temp_links.append(target)

    if dead_or_temp_links:
        raise OfficeExportError(
            'Phát hiện liên kết chart trỏ tới file tạm hoặc file ngoài workbook đích: ' + ', '.join(dead_or_temp_links)
        )
    return {'chart_xml_count': len(chart_parts), 'external_links': external_links}


def build_office_presentation(
    output_pptx_path: Path | str,
    *,
    final_excel_path: Path | str,
    excel_meta: dict[str, Any],
    period: str,
    date_from: str,
    date_to: str,
    hfpa_result: dict[str, Any],
    ftt_result: dict[str, Any],
    preview_dir: Path | str | None = None,
) -> dict[str, Any]:
    """Assemble and verify PPTX via Office/WPS COM Copy/Paste from finalized Excel workbook."""
    import pythoncom
    import win32com.client as client

    pptx_out = Path(output_pptx_path).resolve()
    xlsx_final = Path(final_excel_path).resolve()
    if not xlsx_final.is_file():
        raise OfficeExportError(f'Chưa có workbook Excel nguồn cuối tại {xlsx_final}')
    pptx_out.parent.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    xl = None
    ppt = None
    workbook = None
    presentation = None
    temp_dir = None
    operation = 'khởi tạo WPS/Office'

    try:
        xl = client.DispatchEx('Excel.Application')
        xl.Visible = False
        xl.DisplayAlerts = False
        ppt = client.DispatchEx('PowerPoint.Application')
        ppt.DisplayAlerts = 0
        app_info = _detect_app_label(xl, ppt)

        # Open finalized delivered Excel workbook (NEVER a temp copy!) and evaluate formulas once
        operation = 'mở workbook Excel trong WPS/Office'
        workbook = _retry(lambda: xl.Workbooks.Open(str(xlsx_final), 0, False))
        try:
            xl.CalculateFull()
            workbook.Save()
        except Exception:
            pass

        operation = 'tạo slide và sao chép bảng, biểu đồ từ Excel'
        presentation = ppt.Presentations.Add(WithWindow=False)
        slide_w = 13.333 * 72
        slide_h = 7.5 * 72
        presentation.PageSetup.SlideWidth = slide_w
        presentation.PageSetup.SlideHeight = slide_h

        temp_dir = tempfile.TemporaryDirectory(prefix='qdr-office-', ignore_cleanup_errors=True)
        img_dir = Path(temp_dir.name)
        ranges = excel_meta['ranges']

        def _range_pic(range_key: str, slide: Any, box: tuple[float, float, float, float], fname: str) -> None:
            sheet_name, c1, c2 = ranges[range_key]
            ws = workbook.Worksheets(sheet_name)
            png_path = img_dir / fname
            w_px, h_px = _copy_range_image(ws, f'{c1}:{c2}', png_path)
            _add_fitted_picture(slide, png_path, box, w_px, h_px)

        def _chart_by_title_or_index(ws: Any, title_fragment: str, fallback_idx: int) -> Any | None:
            cnt = int(ws.ChartObjects().Count)
            if cnt == 0:
                return None
            for i in range(1, cnt + 1):
                co = ws.ChartObjects(i)
                try:
                    t = str(co.Chart.ChartTitle.Text or '')
                    if title_fragment.lower() in t.lower():
                        return co
                except Exception:
                    continue
            return ws.ChartObjects(fallback_idx) if fallback_idx <= cnt else None

        expected_charts = 0
        ws_over = workbook.Worksheets('Overview_and_Charts')
        ws_hfpa = workbook.Worksheets('Detail_Breakdown_HFPA')
        ws_ftt = workbook.Worksheets('Detail_Breakdown_FTT')

        # Slide 1: Executive Cover, Status Summary & Dynamic 3-Month HFPA Excel Snapshot (Right-Side Note per AGENTS.md)
        s1 = presentation.Slides.Add(1, 12)
        _add_slide_header(
            s1,
            f'QUARTERLY NIKE DEFECT REVIEW — {period} ({date_from} đến {date_to})',
            f'HFPA NIKE & FTT NIKE độc lập | Danh mục VB: {hfpa_result["catalog_version"]} | Trạng thái DR: CHƯA CHỐT CÔNG THỨC',
            slide_w,
        )
        _range_pic('hfpa_scope_table', s1, (24, 88, 460, 180), 's1_hfpa_scope.png')
        _range_pic('ftt_scope_table', s1, (24, 282, 460, 180), 's1_ftt_scope.png')
        _add_note_box(
            s1,
            (24, 474, 460, 50),
            [
                f'• HFPA Status: {hfpa_result["review"]["status_label"]} | FTT Status: {ftt_result["review"]["status_label"]}',
                f'• DR Notice: {hfpa_result["dr_config"]["note"]}',
            ],
            fill_bgr=0xFEE2E0 if False else 0xFEF8E0,
        )
        _range_pic('hfpa_3month_table', s1, (500, 88, 440, 435), 's1_hfpa_3m.png')

        # Slide 2: HFPA NIKE Overview — Native Chart + Top 3 Groups Table
        s2 = presentation.Slides.Add(2, 12)
        _add_slide_header(
            s2,
            f'HFPA NIKE — TOP 3 NHÓM LỖI CLG & ĐÓNG GÓP THEO KHU ({period})',
            'Phân khu: VH (VH+VH2 trừ VB) | JV (JV+JV2) | VB (Factory Sandals từ danh mục FTT, giữ nguyên số HFPA)',
            slide_w,
        )
        co_h1 = _chart_by_title_or_index(ws_over, 'HFPA NIKE — Top 3 Nhóm', 1)
        if co_h1 is not None:
            _paste_chart(ws_over, co_h1, s2, (24, 86, 456, 290), 'HFPA Top 3 Groups')
            expected_charts += 1
        _range_pic('hfpa_group_table', s2, (490, 86, 450, 265), 's2_hfpa_groups.png')
        _add_note_box(
            s2,
            (24, 392, 916, 130),
            hfpa_result['review']['factual_notes'][:5] or ['(Chưa đủ dữ liệu HFPA NIKE để lập nhận xét số liệu.)'],
        )

        # Slide 3: HFPA NIKE — Detail Availability Status & Top 3 Models (Parallel Dimension)
        s3 = presentation.Slides.Add(3, 12)
        _add_slide_header(
            s3,
            f'HFPA NIKE — TÌNH TRẠNG LỖI CHI TIẾT & TOP 3 MODEL SONG SONG ({period})',
            f'Mức độ hạt nguồn HFPA: {hfpa_result["detail_support_level"]} | Không dùng số FTT để chia nhỏ tổng HFPA',
            slide_w,
        )
        _range_pic('hfpa_detail_matrix', s3, (24, 84, 560, 210), 's3_hfpa_det.png')
        _range_pic('hfpa_model_matrix', s3, (24, 304, 560, 218), 's3_hfpa_mod.png')
        co_hm = _chart_by_title_or_index(ws_over, 'HFPA NIKE — Top Model', 2)
        if co_hm is not None:
            _paste_chart(ws_over, co_hm, s3, (596, 84, 344, 290), 'HFPA Top Models')
            expected_charts += 1

        # Slide 4: FTT NIKE Overview — Native Chart + Top 3 Groups Table
        s4 = presentation.Slides.Add(4, 12)
        _add_slide_header(
            s4,
            f'FTT NIKE — TOP 3 NHÓM LỖI CLG & ĐÓNG GÓP THEO KHU ({period})',
            'Tính độc lập với HFPA | Mẫu số khối lượng: Produced Pairs (A+B+C); số lỗi: Defect Issues across Rework/B/C',
            slide_w,
        )
        co_f1 = _chart_by_title_or_index(ws_over, 'FTT NIKE — Top 3 Nhóm', 3)
        if co_f1 is not None:
            _paste_chart(ws_over, co_f1, s4, (24, 86, 456, 290), 'FTT Top 3 Groups')
            expected_charts += 1
        _range_pic('ftt_group_table', s4, (490, 86, 450, 265), 's4_ftt_groups.png')
        _add_note_box(
            s4,
            (24, 392, 916, 130),
            ftt_result['review']['factual_notes'][:5] or ['(Chưa đủ dữ liệu FTT NIKE để lập nhận xét số liệu.)'],
        )

        # Slide 5: FTT NIKE — Detail / Subtype Breakdown Matrix + Native Chart
        s5 = presentation.Slides.Add(5, 12)
        _add_slide_header(
            s5,
            f'FTT NIKE — PHÂN TÍCH LỖI CHI TIẾT / PHÂN LOẠI TRONG TOP 3 NHÓM ({period})',
            'Mỗi ô hiển thị Top <=3 lỗi thực tế & % trên TỔNG NHÓM tại khu đó (không thu hẹp mẫu số về Top 3)',
            slide_w,
        )
        _range_pic('ftt_detail_matrix', s5, (24, 86, 560, 420), 's5_ftt_det.png')
        co_fd = _chart_by_title_or_index(ws_ftt, 'Top Lỗi Chi Tiết', 1)
        if co_fd is not None:
            _paste_chart(ws_ftt, co_fd, s5, (596, 86, 344, 310), 'FTT Top Details')
            expected_charts += 1

        # Slide 6: FTT NIKE — Parallel Top 3 Models Matrix + Native Model Chart
        s6 = presentation.Slides.Add(6, 12)
        _add_slide_header(
            s6,
            f'FTT NIKE — TOP 3 MODEL TRONG TỪNG NHÓM LỖI (CHIỀU PHÂN TÍCH SONG SONG) ({period})',
            'Model ưu tiên trường Model của FTT NIKE; không gộp GS/PS/TD/WMNS bằng cắt hậu tố',
            slide_w,
        )
        _range_pic('ftt_model_matrix', s6, (24, 86, 560, 420), 's6_ftt_mod.png')
        co_fm = _chart_by_title_or_index(ws_over, 'FTT NIKE — Top Model', 4)
        if co_fm is not None:
            _paste_chart(ws_over, co_fm, s6, (596, 86, 344, 310), 'FTT Top Models')
            expected_charts += 1

        # Slide 7: Review, Completeness & Unconfirmed DR Summary
        s7 = presentation.Slides.Add(7, 12)
        _add_slide_header(
            s7,
            'TRẠNG THÁI DỮ LIỆU, KIỂM TRA ÁNH XẠ & CẤU HÌNH DR CHƯA CHỐT',
            f'Audit trail: {len(hfpa_result["manifests"])} file HFPA NIKE | {len(ftt_result["manifests"])} file FTT NIKE | {hfpa_result["mapping_version"]}',
            slide_w,
        )
        review_lines = [
            f'1. Công thức DR: {hfpa_result["dr_config"]["status"].upper()} — {hfpa_result["dr_config"]["note"]}',
            f'2. HFPA NIKE Status: {hfpa_result["review"]["status_label"]} | Phạm vi thiếu: {hfpa_result["review"]["missing_scopes"] or "Không"}',
            f'3. FTT NIKE Status: {ftt_result["review"]["status_label"]} | Phạm vi thiếu: {ftt_result["review"]["missing_scopes"] or "Không"}',
            f'4. Review items HFPA: {len(hfpa_result["review"]["items"])} | Review items FTT: {len(ftt_result["review"]["items"])}',
        ]
        for item in (hfpa_result['review']['items'] + ftt_result['review']['items'])[:8]:
            review_lines.append(f'   - [{item["metric"]}] {item["status"]}: {item["raw_label"]} ({item["affected_quantity"]:,.0f}) — {item["reason"]}')
        _add_note_box(s7, (24, 92, 912, 410), review_lines)

        slide_count = int(presentation.Slides.Count)
        # Save via native Office/WPS (no XML patching after SaveAs!)
        operation = 'lưu PowerPoint'
        presentation.SaveAs(str(pptx_out))
        operation = 'đóng tài liệu PowerPoint và Excel'
        _close_saved_document(presentation)
        presentation = None
        _close_saved_document(workbook, False)
        workbook = None

        # Close Excel COM completely before reopening PPTX to verify zero dependency on live Excel session
        operation = 'kết thúc phiên Excel/WPS trước khi kiểm tra mở lại'
        old_xl, xl = xl, None
        _quit_after_close(old_xl)
        # WPS can share its COM lifecycle across Spreadsheets and Presentation.
        # Discard both old proxies before creating the verification session.
        old_ppt, ppt = ppt, None
        _quit_after_close(old_ppt)
        old_xl = old_ppt = None

        operation = 'kiểm tra cấu trúc và liên kết PowerPoint'
        pkg_info = _inspect_pptx_package(pptx_out, xlsx_final)

        # Reopen PPTX using real application to verify zero Repair error, exact slides, and live HasChart objects
        operation = 'mở lại PowerPoint bằng kết nối WPS/Office mới'
        ppt, reopened = _open_verification_session(client, pptx_out)
        try:
            operation = 'kiểm tra slide, biểu đồ và xuất ảnh xem trước'
            reopen_slides = int(reopened.Slides.Count)
            reopen_charts = 0
            preview_paths: list[str] = []
            if preview_dir:
                p_dir = Path(preview_dir)
                p_dir.mkdir(parents=True, exist_ok=True)
            for s_idx in range(1, reopen_slides + 1):
                sl = reopened.Slides.Item(s_idx)
                for sh_idx in range(1, int(sl.Shapes.Count) + 1):
                    if getattr(sl.Shapes.Item(sh_idx), 'HasChart', False):
                        reopen_charts += 1
                if preview_dir:
                    png_p = Path(preview_dir) / f'slide_{s_idx}.png'
                    sl.Export(str(png_p.resolve()), 'PNG', 1280, 720)
                    preview_paths.append(str(png_p.resolve()))
            if reopen_slides != slide_count:
                raise OfficeExportError(f'Mở lại PPTX không đủ slide ({reopen_slides} vs {slide_count}).')
            if reopen_charts != expected_charts:
                raise OfficeExportError(f'Mở lại PPTX không đủ chart native ({reopen_charts} vs {expected_charts}).')
        finally:
            # Closing a hidden WPS document can end its process; do not mask a
            # verification error with an expected disconnected cleanup proxy.
            _close_saved_document(reopened)

        return {
            'pptx_path': str(pptx_out),
            'excel_path': str(xlsx_final),
            'slide_count': slide_count,
            'native_charts_pasted': reopen_charts,
            'chart_xml_count': pkg_info['chart_xml_count'],
            'external_links': pkg_info['external_links'],
            'verification_app': app_info['label'],
            'verification_engine': app_info['engine'],
            'reopen_verified': True,
            'slide_previews': preview_paths,
        }
    except OfficeExportError:
        raise
    except Exception as error:
        raise OfficeExportError(f'Lỗi xuất PowerPoint qua {app_info["label"] if "app_info" in locals() else "Office/WPS"} tại thao tác {operation}: {error}') from error
    finally:
        for doc, args in ((presentation, ()), (workbook, (False,))):
            if doc is not None:
                try:
                    if doc is presentation:
                        doc.Saved = True
                    doc.Close(*args)
                except Exception:
                    pass
        if ppt is not None:
            try:
                ppt.Quit()
            except Exception:
                pass
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass
        if temp_dir is not None:
            temp_dir.cleanup()
        pythoncom.CoUninitialize()
