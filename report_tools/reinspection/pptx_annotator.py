"""
Module: pptx_annotator.py
Automates visual inspection overlays (bounding frames, connector lines with dots,
and insight callouts) on Re-Inspection PowerPoint reports.

STANDARD RULES:
- By QC: Threshold is 10%. (Re-Ins / FTT >= 10% for a defect to qualify for framing & connector).
- By MA: Threshold is 30%. (Re-Ins / HFPA >= 30% for a defect to qualify for framing & connector).
- Mapping: Strictly Defect A -> Defect A (same defect name and same segment color).
- Zero overlap: Connectors strictly span inter-column gaps, leaving numbers 100% visible.
"""

import os
import shutil
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

COLOR_RED = RGBColor(255, 0, 0)
COLOR_PURPLE = RGBColor(128, 100, 162)  # #8064A2 matching template
COLOR_BLUE = RGBColor(31, 78, 120)     # Navy/Royal blue matching template
COLOR_GRAY = RGBColor(55, 65, 81)

THRESHOLD_QC = 0.10  # 10% for QC
THRESHOLD_MA = 0.30  # 30% for MA


def add_bounding_box(slide, left, top, width, height, color_rgb, line_width_pt=2.0):
    """Add a transparent rectangle bounding box with colored border."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.background()
    box.line.color.rgb = color_rgb
    box.line.width = Pt(line_width_pt)
    return box


def add_connector_with_oval_ends(slide, x1, y1, x2, y2, color_rgb, line_width_pt=1.2, end_size="sm"):
    """
    Add straight connector line strictly across the inter-column gap,
    with small oval endpoints on the inner edges, ensuring zero overlap with data label numbers.
    """
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    conn.line.color.rgb = color_rgb
    conn.line.width = Pt(line_width_pt)

    ln = conn._element.spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ln')
    if ln is not None:
        headEnd = pptx.oxml.xmlchemy.OxmlElement('a:headEnd')
        headEnd.set('type', 'oval')
        headEnd.set('w', end_size)
        headEnd.set('len', end_size)
        ln.append(headEnd)

        tailEnd = pptx.oxml.xmlchemy.OxmlElement('a:tailEnd')
        tailEnd.set('type', 'oval')
        tailEnd.set('w', end_size)
        tailEnd.set('len', end_size)
        ln.append(tailEnd)

    return conn


def add_insight_callout(slide, left, top, width, height, paragraphs, megaphone_png=None):
    """
    Add concise, executive insight callout matching reference PPTX presentation.
    paragraphs: list of list of (run_text, is_bold, color_rgb, is_italic)
    """
    if megaphone_png and os.path.exists(megaphone_png):
        pic = slide.shapes.add_picture(megaphone_png, left, top + Inches(0.04), Inches(0.36), Inches(0.36))
        pic.name = "Annotator_Megaphone"

    tb_left = left + Inches(0.44)
    tb_width = width - Inches(0.44)
    tb = slide.shapes.add_textbox(tb_left, top, tb_width, height)
    tb.name = "Annotator_Insight"
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    for i, p_spec in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(8)
        p.line_spacing = 1.15
        for r_text, r_bold, r_color, r_italic in p_spec:
            run = p.add_run()
            run.text = r_text
            run.font.name = "Calibri"
            run.font.size = Pt(10.0)
            run.font.bold = r_bold
            run.font.italic = r_italic
            if r_color:
                run.font.color.rgb = r_color
            else:
                run.font.color.rgb = RGBColor(31, 41, 55)

    return tb


def remove_orphan_arrows(slide):
    """Remove leftover unanchored arrow connectors from previous manual drafts."""
    to_remove = []
    for s in slide.shapes:
        if s.name.startswith('Straight Arrow Connector'):
            to_remove.append(s._element)
    for elem in to_remove:
        slide.shapes._spTree.remove(elem)


def remove_previous_annotations(slide):
    """Remove any previously added annotator boxes, connectors, and insight callouts."""
    to_remove = []
    for s in slide.shapes:
        if s.name.startswith("Annotator_") or s.name.startswith("Straight Arrow Connector"):
            to_remove.append(s._element)
    for elem in to_remove:
        slide.shapes._spTree.remove(elem)


def parse_stacked_chart(chart_shape):
    """
    Extract geometry, models, series order and values from OpenXML 100% stacked column chart.
    Returns normalized coordinates (in inches) and segment calculation capabilities.
    """
    root = chart_shape.chart.part._element
    ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
    
    ml = root.find('.//c:plotArea//c:layout//c:manualLayout', ns)
    px = float(ml.find('c:x', ns).attrib['val']) if ml is not None and ml.find('c:x', ns) is not None else 0.0334
    py = float(ml.find('c:y', ns).attrib['val']) if ml is not None and ml.find('c:y', ns) is not None else 0.0532
    pw = float(ml.find('c:w', ns).attrib['val']) if ml is not None and ml.find('c:w', ns) is not None else 0.6651
    ph = float(ml.find('c:h', ns).attrib['val']) if ml is not None and ml.find('c:h', ns) is not None else 0.5019

    sh_l = chart_shape.left / 914400.0
    sh_t = chart_shape.top / 914400.0
    sh_w = chart_shape.width / 914400.0
    sh_h = chart_shape.height / 914400.0

    plot_left = sh_l + sh_w * px
    plot_top = sh_t + sh_h * py
    plot_width = sh_w * pw
    plot_height = sh_h * ph
    plot_bottom = plot_top + plot_height

    lvls = root.findall('.//c:lvl', ns)
    lvl_model = lvls[1] if len(lvls) > 1 else lvls[0]
    models = [p.find('c:v', ns).text for p in lvl_model.findall('c:pt', ns)]
    num_models = max(len(models), 1)
    slot_w = plot_width / float(num_models * 2)

    sers = root.findall('.//c:ser', ns)
    series_data = []
    for s in sers:
        tx_elem = s.find('.//c:tx//c:v', ns)
        dname = tx_elem.text if tx_elem is not None else "Unknown"
        val_pts = {}
        for pt in s.findall('.//c:val//c:pt', ns):
            idx = int(pt.attrib.get('idx', 0))
            ve = pt.find('c:v', ns)
            if ve is not None and ve.text:
                val_pts[idx] = float(ve.text)
        series_data.append((dname, val_pts))

    return {
        'chart_shape': chart_shape,
        'plot_left': plot_left,
        'plot_top': plot_top,
        'plot_width': plot_width,
        'plot_height': plot_height,
        'plot_bottom': plot_bottom,
        'models': models,
        'slot_w': slot_w,
        'series_data': series_data
    }


def calculate_model_defect_segments(cinfo, model_idx):
    """
    Compute exact vertical segments for each defect in Station 1 (col 1) and Station 2 (col 2).
    Returns (segs1, segs2, x1, x2).
    """
    pt1 = model_idx * 2
    pt2 = model_idx * 2 + 1
    tot1 = sum(vals.get(pt1, 0) for _, vals in cinfo['series_data'])
    tot2 = sum(vals.get(pt2, 0) for _, vals in cinfo['series_data'])

    def get_segs(pt, tot):
        segs = {}
        cum = 0.0
        for dname, vals in cinfo['series_data']:
            v = vals.get(pt, 0)
            if v > 0 and tot > 0:
                b_frac = cum / tot
                t_frac = (cum + v) / tot
                cum += v
                c_frac = (b_frac + t_frac) / 2.0
                segs[dname] = {
                    'val': v,
                    'y_center': cinfo['plot_bottom'] - c_frac * cinfo['plot_height'],
                    'b_frac': b_frac,
                    't_frac': t_frac
                }
        return segs

    segs1 = get_segs(pt1, tot1)
    segs2 = get_segs(pt2, tot2)

    # Gap coordinates between col 1 and col 2
    x1 = cinfo['plot_left'] + (pt1 + 0.90) * cinfo['slot_w']
    x2 = cinfo['plot_left'] + (pt1 + 1.10) * cinfo['slot_w']
    return segs1, segs2, x1, x2


def get_model_qualified_defects(cinfo, model_idx, threshold=0.0):
    """
    Find all defects in model_idx where Re-ins / Station1 >= threshold.
    Returns list of dicts with exact defect coordinates for Defect A -> Defect A connector line.
    """
    segs1, segs2, x1, x2 = calculate_model_defect_segments(cinfo, model_idx)
    qualified = []
    for dname, info1 in segs1.items():
        if dname in segs2:
            info2 = segs2[dname]
            v1 = info1['val']
            v2 = info2['val']
            if v1 > 0:
                ratio = v2 / v1
                if ratio >= threshold:
                    qualified.append({
                        'defect': dname,
                        'v1': v1,
                        'v2': v2,
                        'ratio': ratio,
                        'y1': info1['y_center'],
                        'y2': info2['y_center'],
                        'x1': x1,
                        'x2': x2
                    })
    # Sort primarily by ratio descending, secondarily by Re-ins volume
    qualified.sort(key=lambda item: (item['ratio'], item['v2']), reverse=True)
    return qualified


def draw_defect_connector(slide, defect_info, color_rgb, name=None, line_width_pt=1.2):
    """
    Draw connector line strictly connecting Defect A in Station 1 to Defect A in Station 2.
    """
    conn = add_connector_with_oval_ends(
        slide,
        Inches(defect_info['x1']),
        Inches(defect_info['y1']),
        Inches(defect_info['x2']),
        Inches(defect_info['y2']),
        color_rgb,
        line_width_pt=line_width_pt
    )
    if name:
        conn.name = name
    return conn


def draw_model_bounding_box(slide, cinfo, start_model_idx, end_model_idx=None, color_rgb=COLOR_RED, name=None):
    """
    Draw bounding box around a single model or a range of models.
    """
    if end_model_idx is None:
        end_model_idx = start_model_idx
    span_count = end_model_idx - start_model_idx + 1
    
    b_left = cinfo['plot_left'] + (2 * start_model_idx) * cinfo['slot_w'] - 0.03
    b_top = cinfo['plot_top'] - 0.05
    b_width = (2 * span_count) * cinfo['slot_w'] + 0.06
    b_height = cinfo['plot_height'] + 0.20

    box = add_bounding_box(
        slide,
        Inches(b_left),
        Inches(b_top),
        Inches(b_width),
        Inches(b_height),
        color_rgb,
        line_width_pt=2.0
    )
    if name:
        box.name = name
    return box


def annotate_presentation(pptx_path, all_sites_data=None, megaphone_png=None, log_fn=print):
    """
    Main entrypoint called during presentation export.
    Applies dynamic rule-based frames, guide lines (strictly Defect A -> Defect A), and insights to PowerPoint report.
    - By QC: threshold = 10%
    - By MA: threshold = 30%
    """
    if not os.path.exists(pptx_path):
        log_fn(f"[WARNING] Presentation not found for annotation: {pptx_path}")
        return

    if megaphone_png is None or not os.path.exists(megaphone_png):
        curr_dir = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(curr_dir, "assets", "megaphone.png"),
            r"c:\Users\User\Desktop\Re-Inspection\assets\megaphone.png",
            r"c:\Users\User\Desktop\HFPA\assets\megaphone.png",
        ]
        for c in candidates:
            if os.path.exists(c):
                megaphone_png = c
                break

    try:
        prs = pptx.Presentation(pptx_path)
    except Exception as e:
        log_fn(f"[WARNING] Could not open {pptx_path} for annotation: {e}")
        return

    if len(prs.slides) < 4:
        log_fn(f"[WARNING] Presentation has fewer than 4 slides, skipping annotations.")
        return

    # =========================================================================
    # SLIDE 4: JV & JV2
    # =========================================================================
    slide4 = prs.slides[3]
    remove_previous_annotations(slide4)
    remove_orphan_arrows(slide4)

    charts_s4 = [s for s in slide4.shapes if s.has_chart]
    charts_s4.sort(key=lambda s: s.top)

    red_defects_summary = []
    purple_defects_summary = []

    if len(charts_s4) >= 4:
        c_jv_qc  = parse_stacked_chart(charts_s4[0])
        c_jv_ma  = parse_stacked_chart(charts_s4[1])
        c_jv2_qc = parse_stacked_chart(charts_s4[2])
        c_jv2_ma = parse_stacked_chart(charts_s4[3])

        # ---------------------------------------------------------------------
        # 1. JV BY QC: Find qualified model with highest violation (threshold = 10%)
        # ---------------------------------------------------------------------
        best_qc_model = None
        best_qc_defect = None
        max_qc_ratio = -1.0

        for m_i in range(len(c_jv_qc['models'])):
            q_defs = get_model_qualified_defects(c_jv_qc, m_i, THRESHOLD_QC)
            if q_defs:
                top_d = q_defs[0]
                if top_d['ratio'] > max_qc_ratio:
                    max_qc_ratio = top_d['ratio']
                    best_qc_model = m_i
                    best_qc_defect = top_d

        if best_qc_model is None:
            # Fallback to model with highest repeating defect volume if none exceeds 10%
            for m_i in range(len(c_jv_qc['models'])):
                q_defs = get_model_qualified_defects(c_jv_qc, m_i, threshold=0.001)
                if q_defs:
                    top_d = q_defs[0]
                    if top_d['v2'] > (best_qc_defect['v2'] if best_qc_defect else -1):
                        best_qc_model = m_i
                        best_qc_defect = top_d

        if best_qc_model is not None and best_qc_defect is not None:
            draw_model_bounding_box(slide4, c_jv_qc, best_qc_model, color_rgb=COLOR_RED, name="Annotator_Box_JV_QC")
            draw_defect_connector(slide4, best_qc_defect, COLOR_RED, name="Annotator_Connector_JV_QC", line_width_pt=1.2)
            red_defects_summary.append((c_jv_qc['models'][best_qc_model], best_qc_defect))
            log_fn(f"  [JV QC] Framed Model {best_qc_model} ({c_jv_qc['models'][best_qc_model]}): {best_qc_defect['defect']} ({best_qc_defect['ratio']:.1%})")

        # ---------------------------------------------------------------------
        # 2. JV2 BY MA: Red Frame on Model 0 (threshold = 30%)
        # ---------------------------------------------------------------------
        q_defs_m0 = get_model_qualified_defects(c_jv2_ma, 0, THRESHOLD_MA)
        if not q_defs_m0:
            q_defs_m0 = get_model_qualified_defects(c_jv2_ma, 0, threshold=0.001)

        if q_defs_m0:
            top_m0_def = q_defs_m0[0]
            draw_model_bounding_box(slide4, c_jv2_ma, 0, color_rgb=COLOR_RED, name="Annotator_Box_JV2_MA_Red")
            draw_defect_connector(slide4, top_m0_def, COLOR_RED, name="Annotator_Connector_JV2_MA_Red", line_width_pt=1.2)
            red_defects_summary.append((c_jv2_ma['models'][0], top_m0_def))
            log_fn(f"  [JV2 MA Red] Framed Model 0 ({c_jv2_ma['models'][0]}): {top_m0_def['defect']} ({top_m0_def['ratio']:.1%})")

        # ---------------------------------------------------------------------
        # 3. JV2 BY MA: Purple Frame spanning Models 1 to 4
        # ---------------------------------------------------------------------
        purple_idx = 1
        for m_i in range(1, min(5, len(c_jv2_ma['models']))):
            q_defs_mi = get_model_qualified_defects(c_jv2_ma, m_i, THRESHOLD_MA)
            if q_defs_mi:
                top_def = q_defs_mi[0]
                draw_defect_connector(slide4, top_def, COLOR_PURPLE, name=f"Annotator_Connector_JV2_MA_Purple_{purple_idx}", line_width_pt=1.2)
                purple_idx += 1
                purple_defects_summary.append((c_jv2_ma['models'][m_i], top_def))
                log_fn(f"  [JV2 MA Purple] Model {m_i} ({c_jv2_ma['models'][m_i]}): {top_def['defect']} ({top_def['ratio']:.1%})")

        if purple_defects_summary:
            draw_model_bounding_box(
                slide4, c_jv2_ma, 1, min(4, len(c_jv2_ma['models']) - 1),
                color_rgb=COLOR_PURPLE, name="Annotator_Box_JV2_MA_Purple")

    paragraphs_slide4 = []
    for label, color, defects in [
        ("RED FRAME", COLOR_RED, red_defects_summary),
        ("PURPLE FRAME", COLOR_PURPLE, purple_defects_summary),
    ]:
        if not defects:
            continue
        by_trend = {}
        for model, comparison in defects:
            before, after = comparison['v1'], comparison['v2']
            trend = (
                "decreased but were still found" if after < before else
                "increased" if after > before else "stayed at the same level"
            )
            by_trend.setdefault(trend, {})[f"{comparison['defect']} on {model}"] = None

        paragraph = [(label, True, color, False), (". ", False, None, False)]
        for trend, names in by_trend.items():
            paragraph.extend([
                (", ".join(names), True, None, False),
                (f" {trend} at re-inspection. ", False, None, False),
            ])
        paragraphs_slide4.append(paragraph)

    if paragraphs_slide4:
        tb_insight = add_insight_callout(
            slide4, Inches(8.85), Inches(3.20), Inches(4.35), Inches(3.20),
            paragraphs_slide4, megaphone_png=megaphone_png)
        tb_insight.name = "Annotator_Insight_JV"

    # =========================================================================
    # SLIDE 3: VH & VH2
    # =========================================================================
    slide3 = prs.slides[2]
    remove_previous_annotations(slide3)
    remove_orphan_arrows(slide3)

    charts_s3 = [s for s in slide3.shapes if s.has_chart]
    charts_s3.sort(key=lambda s: s.top)
    if len(charts_s3) >= 4:
        c_vh_qc  = parse_stacked_chart(charts_s3[0])
        c_vh_ma  = parse_stacked_chart(charts_s3[1])
        c_vh2_qc = parse_stacked_chart(charts_s3[2])
        c_vh2_ma = parse_stacked_chart(charts_s3[3])

        for cinfo, thresh, site_label in [
            (c_vh_qc, THRESHOLD_QC, "VH QC"),
            (c_vh_ma, THRESHOLD_MA, "VH MA"),
            (c_vh2_qc, THRESHOLD_QC, "VH2 QC"),
            (c_vh2_ma, THRESHOLD_MA, "VH2 MA")
        ]:
            for m_i in range(len(cinfo['models'])):
                q_defs = get_model_qualified_defects(cinfo, m_i, thresh)
                if q_defs:
                    top_d = q_defs[0]
                    draw_model_bounding_box(slide3, cinfo, m_i, color_rgb=COLOR_RED, name=f"Annotator_Box_{site_label}_{m_i}")
                    draw_defect_connector(slide3, top_d, COLOR_RED, name=f"Annotator_Connector_{site_label}_{m_i}", line_width_pt=1.2)
                    log_fn(f"  [{site_label}] Framed Model {m_i} ({cinfo['models'][m_i]}): {top_d['defect']} ({top_d['ratio']:.1%})")

    try:
        prs.save(pptx_path)
        log_fn(f"[OK] Visual inspection annotations (frames, connectors & insights) applied to: {pptx_path}")
    except PermissionError:
        base, ext = os.path.splitext(pptx_path)
        alt_path = f"{base}_updated{ext}"
        prs.save(alt_path)
        log_fn(f"[WARNING] {pptx_path} is locked. Saved annotated presentation to: {alt_path}")


def annotate_september_report(pptx_path, megaphone_png=None):
    """Backward-compatible alias for annotate_presentation."""
    return annotate_presentation(pptx_path, megaphone_png=megaphone_png)


def main():
    target_pptx = r"c:\Users\User\Desktop\Re-Inspection\Output\Re-Ins Report-September-2026.pptx"
    annotate_presentation(target_pptx)


if __name__ == "__main__":
    main()
