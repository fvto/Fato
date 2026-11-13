#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_copq_pptx.py
===================
Clone the COPQ/FTT PowerPoint report and update its data for a new reporting
period using the existing template layout.

SAFETY (hard):
  * Original .pptx is NEVER modified (we copy it first).
  * COPQ chart categories & series are KEPT; only their numeric values
    (and, where the period changes them, category labels) are rewritten by a
    (series x category) lookup.
  * Missing COPQ data -> OLD value kept + WARNING logged (never fabricated).
  * The four FTT defect charts and their cell-based legends are copied from
    FTT_Combined_Report.xlsx through Office, preserving all source defects.
    They replace the corresponding template slots. No FTT chart is rebuilt
    from raw data or recolored in PowerPoint.

DATA CONTRACT (verified read-only):
  * COPQ_Clean.xlsx : one sheet per site (VH,VH2,VH3,VH4,JV,JV2,JV3). JV3->JVB.
    Columns: Date, Category, Type, Factory, Fty/Line, Model,
    Defective Qty(Pair), Ttl cost ($), Remark (Re-inspection Station).
  * FTT_Combined_Report.xlsx : Site_VH/VH2/JV/JV2, ColorMap_bc,
    ColorMap_BC New 2026. Site_* cols: Site,Date,Model,Metric,Defect_Issue,FTT_Qty.

Run:  python3 update_copq_pptx.py
"""

import os
import sys
import shutil
import logging
import math
from lxml import etree

import pandas as pd
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(HERE) in ("Output", "Source"):
    BASE = os.path.dirname(HERE)
    OUT_DIR = os.path.join(BASE, "Output")
else:
    BASE = HERE
    OUT_DIR = os.path.join(HERE, "Output")

for p in (BASE, os.path.join(BASE, "Source")):
    if p not in sys.path:
        sys.path.insert(0, p)

from monthly_reporting import classify_copq_type
from copq_office_charts import copy_ftt_charts

SRC_PPTX  = os.path.join(BASE, "Template_COPQ", "Template_COPQ.pptx")
COPQ_XLSX = os.path.join(OUT_DIR, "COPQ_Clean.xlsx")
FTT_XLSX  = os.path.join(OUT_DIR, "FTT_Combined_Report.xlsx")

def detect_report_period(clean_path):
    """Detect reporting month from COPQ_Clean.xlsx dates."""
    try:
        import datetime as dt
        if os.path.exists(clean_path):
            df = pd.read_excel(clean_path, sheet_name="COPQ_Clean", nrows=500)
            if "Date" in df.columns:
                dates = pd.to_datetime(df["Date"], errors="coerce").dropna()
                if not dates.empty:
                    max_d = dates.max()
                    p_label = max_d.strftime("%b, %Y")
                    d_month = max_d.strftime("%b-%Y")
                    out_f = f"COPQ_Report_{max_d.strftime('%b_%Y')}.pptx"
                    return p_label, d_month, out_f
    except Exception:
        pass
    return "Sep, 2026", "Sep-2026", "COPQ_Report_Sep_2026.pptx"

_detected_period, _detected_db_month, _detected_out = detect_report_period(COPQ_XLSX)

# --- database (accumulated source of truth) ---
DB_OVERVIEW = os.path.join(BASE, "Database", "CoPQ database 25.xlsx")
DB_TYPE     = os.path.join(BASE, "Database", "CoPQ_type_analysis.xlsx")
DB_MONTH    = _detected_db_month
PERIOD_LABEL = _detected_period
OUT_PPTX    = os.path.join(OUT_DIR, _detected_out)

ANALYSIS_CHART_LAYOUT = {
    2: {
        "model": [("Chart 23", 110874, 2011680, 1920240, 2286000), ("Chart 31", 2116666, 2011680, 1911245, 2286000), ("Chart 34", 4134853, 2011680, 1920240, 2286000), ("Chart 36", 6128462, 2011680, 1922145, 2285365), ("Chart 17", 8186988, 2011680, 1922145, 2285365), ("Chart 27", 10196144, 2011680, 1920240, 2286000)],
        "pie": [("Chart 22", 108727, 4754880, 1920240, 2011680), ("Chart 32", 2092533, 4754880, 1919605, 2011045), ("Chart 33", 4126898, 4754880, 1919605, 2011680), ("Chart 37", 6140024, 4782312, 1920240, 2011680), ("Chart 19", 8205295, 4736592, 1920240, 2011045), ("Chart 21", 10187970, 4754880, 1918970, 1982470)],
    },
    3: {
        "model": [("Chart 4", 115503, 2103120, 2598821, 2280920), ("Chart 13", 3208235, 2135019, 2578691, 2281555), ("Chart 15", 6198669, 2103120, 2551801, 2324452), ("Chart 17", 9211377, 2103120, 2541069, 2223346)],
        "pie": [("Chart 9", 335280, 4920000, 1828800, 1900000), ("Chart 11", 3616960, 4920000, 1828800, 1900000), ("Chart 19", 6664960, 4920000, 1828800, 1900000), ("Chart 21", 9540240, 4920000, 1869440, 1900000)],
    },
}

DECK_SITES = ["VH", "VH2", "VH3", "VH4", "JV", "JV2", "JVB"]
PIE_COLORS = ["3366FF", "33CCFF", "BFBFBF", "99CCFF", "7F7F7F"]
DATA_SITE_ALIAS = {"JV3": "JVB"}    # data sheet -> deck site

# COPQ Type values to DROP entirely from the report (totals, per-site bars,
# monthly charts, factory pies, top-3 model bars). Set to [] to include all.
# Per your instruction: Rework is excluded so the report covers only
# Touch-up / Re-inspection / B/C Grade.
EXCLUDE_TYPES = ["Rework"]

def map_bar_cat(row):
    """Map one COPQ_Clean row to a deck per-site bar category, per the SOP.
    Returns 'Touch up paint', 'Re-inspection', 'B/C Grade', or None (dropped)."""
    group = classify_copq_type(row)
    if group == "Touch-up":
        return "Touch up paint"
    if group == "Reinspection":
        return "Re-inspection"
    if group == "B/C":
        return "B/C Grade"
    return None

def non_rework(df):
    """Return a copy of the per-site dataframe with EXCLUDE_TYPES dropped,
    so totals/bars/monthly/factory/model charts never include the excluded
    types (e.g. Rework)."""
    if df is None or df.empty:
        return df
    mask = ~df["Type"].astype(str).str.strip().isin(EXCLUDE_TYPES)
    return df[mask].copy()

SLIDE1_GROUPS = [                   # top-to-bottom site groups on slide 1
    {"site": "VH",  "model_bar": "Chart 24", "factory_pie": "Chart 26", "defect": "Chart 78"},
    {"site": "VH2", "model_bar": "Chart 35", "factory_pie": "Chart 28", "defect": "Chart 80"},
    {"site": "JV",  "model_bar": "Chart 38", "factory_pie": "Chart 30", "defect": "Chart 92"},
    {"site": "JV2", "model_bar": "Chart 40", "factory_pie": "Chart 33", "defect": "Chart 96"},
]

MONTHLY_CHARTS = {"VH": "VH", "VH2": "VH2", "JV": "JV", "JV2": "JV2"}

DEFECT_VALUE_MODE = "proportion"    # 'proportion' (cols sum to 1) or 'count'

# Defect types that appear in the new data but are NOT in any color palette
# (deck embedded colors, ColorMap_bc, or ColorMap_BC New 2026). To add a new
# defect series to the Top-3 Defect charts, assign it a hex here, e.g.
#   "C/Stitching margin": "FFC000",
# The script STOPS and lists any missing defect here so you make the call,
# never auto-inventing a color.
MISSING_DEFECT_COLORS = {
    # "C/Stitching margin": "FFC000",   # <-- uncomment & set once you decide
}

# Slide-0 per-site bar charts in deck table-column order VH..JVB
SITE_BAR_CHARTS = ["Chart 18", "Chart 20", "Chart 22", "Chart 24",
                   "Chart 26", "Chart 28", "Chart 30"]

# --------------------------------------------------------------------------
# Logging
# --------------------------------------------------------------------------
log = logging.getLogger("copq")
log.setLevel(logging.INFO)
log.addHandler(logging.StreamHandler(sys.stdout))
WARN = []

# --------------------------------------------------------------------------
# Data loading (read-only)
# --------------------------------------------------------------------------
def load_copq():
    xl = pd.ExcelFile(COPQ_XLSX)
    out = {}
    for sheet in xl.sheet_names:
        if sheet == "COPQ_Clean":
            continue
        df = xl.parse(sheet, header=1)
        if "Date" not in df.columns and len(df) and str(df.iloc[0, 0]).startswith("Cost Of Poor Quality"):
            df = xl.parse(sheet, header=2)
        df = df[df.get("Date", pd.Series([None])).notna()]
        if EXCLUDE_TYPES:
            before = len(df)
            mask = ~df["Type"].astype(str).str.strip().isin(EXCLUDE_TYPES)
            df = df[mask]
            log.info(f"COPQ '{sheet}': excluded {before - len(df)} rows of "
                     f"{EXCLUDE_TYPES} (kept {len(df)})")
        for c in ("Ttl cost ($)", "Defective Qty(Pair)"):
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
        deck_site = DATA_SITE_ALIAS.get(sheet, sheet)
        out[deck_site] = df
        log.info(f"COPQ loaded: '{sheet}' -> deck '{deck_site}', {len(df)} rows, "
                 f"cost=${df['Ttl cost ($)'].sum():,.2f}")
    return out

def load_ftt():
    if not os.path.exists(FTT_XLSX):
        log.warning(f"FTT report file not found at {FTT_XLSX}; defect charts will retain template data.")
        return {}, {}
    xl = pd.ExcelFile(FTT_XLSX)
    site_def = {}
    for s in ("Site_VH", "Site_VH2", "Site_JV", "Site_JV2"):
        if s in xl.sheet_names:
            d = xl.parse(s)
            d["FTT_Qty"] = pd.to_numeric(d.get("FTT_Qty"), errors="coerce").fillna(0)
            site_def[s.replace("Site_", "")] = d
    cmap = {}
    # ColorMap_BC New 2026 : header row 3 (0-idx); defect in col2, HEX col7, RGB 4/5/6
    if "ColorMap_BC New 2026" in xl.sheet_names:
        d = xl.parse("ColorMap_BC New 2026", header=None)
        for _, r in d.iterrows():
            name = r[2]
            if not isinstance(name, str) or name.strip() in ("", "Defect type"):
                continue
            name = norm_defect(name)
            hexv = r[7] if pd.notna(r[7]) else None
            if hexv and str(hexv).strip():
                cmap[name] = str(hexv).strip().lstrip("#")
            else:
                try:
                    cmap[name] = f"{int(r[4]):02X}{int(r[5]):02X}{int(r[6]):02X}"
                except Exception:
                    pass
    # ColorMap_bc : header row 0; defect col0, HEX col4, RGB 1/2/3
    if "ColorMap_bc" in xl.sheet_names:
        d = xl.parse("ColorMap_bc", header=0)
        for _, r in d.iterrows():
            name = r.get("Defect type")
            if not isinstance(name, str) or not name.strip():
                continue
            name = norm_defect(name)
            hexv = r.get("HEX")
            if pd.notna(hexv) and str(hexv).strip():
                cmap.setdefault(name, str(hexv).strip().lstrip("#"))
            else:
                try:
                    cmap.setdefault(name, f"{int(r['Red']):02X}{int(r['Green']):02X}{int(r['Blue']):02X}")
                except Exception:
                    pass
    log.info(f"FTT loaded: site defect sheets={list(site_def)}, "
             f"colormap={len(cmap)} entries")
    return site_def, cmap

def build_palette(prs, cmap):
    """Palette: deck-embedded series colors first (authoritative for consistency),
    then color-map file. Only C/*/B/* defect names are collected."""
    pal = {}
    for slide in prs.slides:
        for shp in slide.shapes:
            if not shp.has_chart:
                continue
            for ser in _ser_elements(shp.chart):
                nm = _ser_name(ser)
                if nm and (nm.startswith("C/") or nm.startswith("B/")):
                    col = _ser_solid_color(ser)
                    if col:
                        pal[norm_defect(nm)] = col
    for k, v in cmap.items():
        pal.setdefault(norm_defect(k), v)
    interior_key = norm_defect("C/ Interior")
    interior_defect_key = norm_defect("C/Interior defect")
    if interior_key in pal:
        pal[interior_defect_key] = pal[interior_key]
    for k, v in MISSING_DEFECT_COLORS.items():
        pal[norm_defect(k)] = v
    log.info(f"PALETTE built: {len(pal)} defect colors "
             f"(deck={sum(1 for k in pal if k in _deck_defects(prs))})")
    return pal

def norm_defect(name):
    """Normalize defect labels so palette/document/data keys match despite
    cosmetic differences: trim, lowercase, collapse whitespace, remove the
    stray space after the slash, and drop hyphens (e.g. 'C/ Stitching margin',
    'C/Airbag defect' vs 'C/ Airbag Defect', 'C/Color mis-match' vs
    'C/Color mismatch' all collapse to the same key)."""
    if not isinstance(name, str):
        return name
    s = name.strip().lower()
    s = s.replace("-", "")
    s = s.replace(" /", "/").replace("/ ", "/")
    s = " ".join(s.split())
    return s

def _deck_defects(prs):
    s = set()
    for slide in prs.slides:
        for shp in slide.shapes:
            if not shp.has_chart:
                continue
            for ser in _ser_elements(shp.chart):
                nm = _ser_name(ser)
                if nm and (nm.startswith("C/") or nm.startswith("B/")):
                    s.add(norm_defect(nm))
    return s

# --------------------------------------------------------------------------
# OOXML helpers
# --------------------------------------------------------------------------
def _ser_elements(chart):
    return list(chart._chartSpace.chart.plotArea.iter(qn('c:ser')))

def _ser_name(ser):
    tx = ser.find(qn('c:tx'))
    if tx is None:
        return None
    strref = tx.find(qn('c:strRef'))
    if strref is None:
        return None
    cache = strref.find(qn('c:strCache'))
    if cache is None:
        return None
    pt = cache.find(qn('c:pt'))
    if pt is None:
        return None
    v = pt.find(qn('c:v'))
    return v.text if v is not None else None

def _ser_solid_color(ser):
    sp = ser.find(qn('c:spPr'))
    if sp is None:
        return None
    sf = sp.find(qn('a:solidFill'))
    if sf is None:
        return None
    srgb = sf.find(qn('a:srgbClr'))
    if srgb is not None:
        return srgb.get('val')
    return None

def _cat_strs(ser):
    cat = ser.find(qn('c:cat'))
    if cat is None:
        return []
    strref = cat.find(qn('c:strRef'))
    if strref is not None:
        cache = strref.find(qn('c:strCache'))
        if cache is not None:
            return [pt.find(qn('c:v')).text for pt in cache.findall(qn('c:pt'))]
    numref = cat.find(qn('c:numRef'))
    if numref is not None:
        cache = numref.find(qn('c:numCache'))
        if cache is not None:
            return [pt.find(qn('c:v')).text for pt in cache.findall(qn('c:pt'))]
    return []

def set_series_values_ordered(ser, values, format_code=None):
    val = ser.find(qn('c:val'))
    if val is None:
        return
    nref = val.find(qn('c:numRef'))
    if nref is None:
        nref = etree.SubElement(val, qn('c:numRef'))
    vcache = nref.find(qn('c:numCache'))
    if vcache is None:
        vcache = etree.SubElement(nref, qn('c:numCache'))
    if format_code:
        fc = vcache.find(qn('c:formatCode'))
        if fc is None:
            fc = etree.SubElement(vcache, qn('c:formatCode'))
        fc.text = format_code
    for pt in vcache.findall(qn('c:pt')):
        vcache.remove(pt)
    pc = vcache.find(qn('c:ptCount'))
    if pc is None:
        pc = etree.SubElement(vcache, qn('c:ptCount'))
    pc.set('val', str(len(values)))
    for i, v in enumerate(values):
        if v is None:
            continue
        pt = etree.SubElement(vcache, qn('c:pt'))
        pt.set('idx', str(i))
        if format_code:
            pt.set(qn('c:formatCode'), format_code)
        ve = etree.SubElement(pt, qn('c:v'))
        ve.text = f"{v:.4f}".rstrip("0").rstrip(".") if isinstance(v, float) else str(v)

def set_series_categories(ser, cats):
    cat = ser.find(qn('c:cat'))
    if cat is None:
        return
    strref = cat.find(qn('c:strRef'))
    if strref is None:
        return
    cache = strref.find(qn('c:strCache'))
    if cache is None:
        cache = etree.SubElement(strref, qn('c:strCache'))
    for pt in cache.findall(qn('c:pt')):
        cache.remove(pt)
    pc = cache.find(qn('c:ptCount'))
    if pc is None:
        pc = etree.SubElement(cache, qn('c:ptCount'))
    pc.set('val', str(len(cats)))
    for i, c in enumerate(cats):
        pt = etree.SubElement(cache, qn('c:pt'))
        pt.set('idx', str(i))
        v = etree.SubElement(pt, qn('c:v'))
        v.text = str(c)

def strip_filtered_extensions(chart):
    """Strip Microsoft Office 2013+ filtered extensions (c:extLst) so PowerPoint displays newly updated live data without falling back to stale filtered cache."""
    plot_area = chart._chartSpace.chart.plotArea
    for ext in list(plot_area.iter(qn('c:extLst'))):
        p = ext.getparent()
        if p is not None:
            p.remove(ext)

def format_model_chart(chart, rot=-35, font_sz=650):
    """Keep model chart labels readable after replacing template cache data."""
    plot_area = chart._chartSpace.chart.plotArea

    # Ensure manualLayout leaves enough room at bottom for rotated category labels
    ml = plot_area.find('.//' + qn('c:manualLayout'))
    if ml is not None:
        y_el = ml.find(qn('c:y'))
        h_el = ml.find(qn('c:h'))
        if y_el is not None and h_el is not None:
            try:
                y_val = float(y_el.attrib.get('val', '0.2'))
                h_val = float(h_el.attrib.get('val', '0.58'))
                if y_val + h_val > 0.80:
                    y_el.set('val', '0.18')
                    h_el.set('val', '0.58')
            except Exception:
                pass

    for axis in list(plot_area.iter(qn('c:catAx'))):
        tx_pr = axis.find(qn('c:txPr'))
        if tx_pr is None:
            tx_pr = etree.SubElement(axis, qn('c:txPr'))
        body_pr = tx_pr.find(qn('a:bodyPr'))
        if body_pr is None:
            body_pr = etree.SubElement(tx_pr, qn('a:bodyPr'))
        body_pr.set('wrap', 'none')
        body_pr.set('rot', str(int(rot * 60000)))
        body_pr.set('vertOverflow', 'clip')
        for latin in tx_pr.iter(qn('a:latin')):
            latin.set('typeface', 'Arial')
        for def_rpr in tx_pr.iter(qn('a:defRPr')):
            def_rpr.set('sz', str(font_sz))

    for axis in list(plot_area.iter(qn('c:valAx'))):
        tx_pr = axis.find(qn('c:txPr'))
        if tx_pr is not None:
            for def_rpr in tx_pr.iter(qn('a:defRPr')):
                def_rpr.set('sz', '600')

    for ser in _ser_elements(chart):
        d_lbl_groups = ser.findall(qn('c:dLbls'))
        if not d_lbl_groups:
            continue
        for d_lbls in d_lbl_groups:
            pos = d_lbls.find(qn('c:dLblPos'))
            if pos is None:
                pos = etree.SubElement(d_lbls, qn('c:dLblPos'))
            pos.set('val', 'outEnd')
            num_fmts = list(d_lbls.iter(qn('c:numFmt')))
            if not num_fmts:
                num_fmts = [etree.SubElement(d_lbls, qn('c:numFmt'))]
            for num_fmt in num_fmts:
                num_fmt.set('formatCode', '"$"#,##0.00')
                num_fmt.set('sourceLinked', '0')
            tx_prs = list(d_lbls.iter(qn('c:txPr')))
            if not tx_prs:
                tx_prs = [etree.SubElement(d_lbls, qn('c:txPr'))]
            for tx_pr in tx_prs:
                for body_pr in tx_pr.iter(qn('a:bodyPr')):
                    body_pr.set('wrap', 'none')
                for latin in tx_pr.iter(qn('a:latin')):
                    latin.set('typeface', 'Arial')
                for def_rpr in tx_pr.iter(qn('a:defRPr')):
                    def_rpr.set('sz', '700')
                    def_rpr.set('b', '1')
    strip_filtered_extensions(chart)

def set_series_color(ser, hex_color):
    sp = ser.find(qn('c:spPr'))
    if sp is None:
        return
    sf = sp.find(qn('a:solidFill'))
    if sf is None:
        sf = etree.SubElement(sp, qn('a:solidFill'))
    for bad in sf.findall(qn('a:schemeClr')):
        sf.remove(bad)
    srgb = sf.find(qn('a:srgbClr'))
    if srgb is None:
        srgb = etree.SubElement(sf, qn('a:srgbClr'))
    srgb.set('val', hex_color)

def set_series_values_by_category(ser, cat_to_val, default=0.0, warn=True):
    cats = _cat_strs(ser)
    vals = []
    for c in cats:
        if c is None:
            vals.append(default)
        elif c in cat_to_val:
            vals.append(float(cat_to_val[c]))
        else:
            vals.append(default)
            if warn and c != "":
                WARN.append(f"series '{_ser_name(ser)}' cat '{c}' not in data -> 0")
    set_series_values_ordered(ser, vals)

def get_chart_by_name(slide, name):
    for shp in slide.shapes:
        if shp.has_chart and shp.name == name:
            return shp.chart
    return None

def apply_analysis_layout(prs):
    """Apply approved analysis-slide geometry without touching chart data."""
    for slide_index, groups in ANALYSIS_CHART_LAYOUT.items():
        slide = prs.slides[slide_index]
        for group in groups.values():
            for name, left, top, width, height in group:
                shp = next((s for s in slide.shapes if s.name == name), None)
                if shp is not None:
                    shp.left, shp.top, shp.width, shp.height = left, top, width, height

# --------------------------------------------------------------------------
# Updaters
# --------------------------------------------------------------------------
def clean_model_name(name, max_len=22):
    """Clean model name: collapse whitespace, trim, limit length cleanly."""
    if not isinstance(name, str) or str(name).strip().lower() in ("none", "nan", ""):
        return ""
    s = " ".join(name.strip().split())
    # Smart strip common brand prefixes but preserve gender qualifiers (WMNS, W)
    if s.upper().startswith("WMNS NIKE "):
        s = "WMNS " + s[10:].strip()
    elif s.upper().startswith("W NIKE "):
        s = "W " + s[7:].strip()
    elif s.upper().startswith("NIKE SB "):
        s = s[8:].strip()
    elif s.upper().startswith("NIKE "):
        s = s[5:].strip()
    if len(s) > max_len:
        return s[:max_len-2].strip() + ".."
    return s

def update_site_bar(prs, copq):
    """Update Slide 0 per-site bar charts (Chart 18, 20, 22, 24, 26, 28, 30).
    Horizontal bar chart category order:
      Index 0: Touch up paint (bottom)
      Index 1: Re-inspection (middle)
      Index 2: B/C Grade (top)
    """
    slide0 = list(prs.slides)[0]
    for name, site in zip(SITE_BAR_CHARTS, DECK_SITES):
        ch = get_chart_by_name(slide0, name)
        if ch is None:
            WARN.append(f"site bar '{name}' not found"); continue
        df = copq.get(site)
        cat_sum = {"Touch up paint": 0.0, "Re-inspection": 0.0, "B/C Grade": 0.0}
        if df is not None and not df.empty:
            for _, r in df.iterrows():
                mapped = map_bar_cat(r)
                if mapped:
                    cat_sum[mapped] += float(r.get("Ttl cost ($)", 0) or 0)
        vals = [
            cat_sum["Touch up paint"] if cat_sum["Touch up paint"] > 0 else None,
            cat_sum["Re-inspection"] if cat_sum["Re-inspection"] > 0 else None,
            cat_sum["B/C Grade"] if cat_sum["B/C Grade"] > 0 else None,
        ]
        for ser in _ser_elements(ch):
            set_series_values_ordered(ser, vals, format_code='"$"#,##0.00')

        # Remove leader lines and manual layout from data labels so no stray lines are drawn
        for dl in ch._chartSpace.chart.plotArea.iter(qn('c:dLbls')):
            for elem in list(dl.iter()):
                if elem.tag.endswith('showLeaderLines'):
                    elem.set('val', '0')
                if elem.tag.endswith('leaderLines'):
                    p = elem.getparent()
                    if p is not None:
                        p.remove(elem)
            for dlbl in dl.findall(qn('c:dLbl')):
                ml = dlbl.find(qn('c:layout'))
                if ml is not None:
                    dlbl.remove(ml)

        # Ensure axis max is large enough so bar labels don't get forced inside the bar
        valid_vals = [v for v in vals if v is not None]
        max_val = max(valid_vals) if valid_vals else 0.0
        val_ax = ch._chartSpace.chart.plotArea.find(qn('c:valAx'))
        if val_ax is not None:
            scaling = val_ax.find(qn('c:scaling'))
            if scaling is not None:
                max_elem = scaling.find(qn('c:max'))
                if max_elem is not None:
                    cur_max = float(max_elem.get('val', 9000))
                    if max_val > cur_max * 0.65:
                        new_max = max(cur_max, (int(max_val * 1.55) // 1000 + 1) * 1000)
                        max_elem.set('val', str(int(new_max)))

        log.info(f"site bar '{name}' ({site}): TU=${cat_sum['Touch up paint']:,.0f}, "
                 f"RE=${cat_sum['Re-inspection']:,.0f}, BC=${cat_sum['B/C Grade']:,.0f}")

def update_na_badges(prs, copq):
    """Position and display N/A badges on Slide 0 exactly where a category has $0 cost.
    Park unused badges offscreen so no stale N/A badges overlap real data."""
    slide0 = list(prs.slides)[0]
    SITE_X = {
        "VH":  2614930,
        "VH2": 3904933,
        "VH3": 5300028,
        "VH4": 6634163,
        "JV":  7976235,
        "JV2": 9343390,
        "JVB": 10678160,
    }
    CAT_Y = {
        "B/C Grade": 2919675,
        "Re-inspection": 3622620,
        "Touch up paint": 4264605,
    }
    needed_slots = []
    for site in DECK_SITES:
        df = copq.get(site)
        cat_sum = {"Touch up paint": 0.0, "Re-inspection": 0.0, "B/C Grade": 0.0}
        if df is not None and not df.empty:
            for _, r in df.iterrows():
                m = map_bar_cat(r)
                if m:
                    cat_sum[m] += float(r.get("Ttl cost ($)", 0) or 0)
        for cat in ["B/C Grade", "Re-inspection", "Touch up paint"]:
            if cat_sum[cat] < 0.01:
                needed_slots.append((SITE_X[site], CAT_Y[cat]))

    na_shapes = [s for s in slide0.shapes if s.name == "Rectangle 64"]
    while len(na_shapes) < len(needed_slots):
        new_shp = slide0.shapes.add_shape(1, 0, 0, 383540, 245110)
        new_shp.name = "Rectangle 64"
        new_shp.fill.background()
        new_shp.line.color.rgb = RGBColor(220, 220, 220)
        new_shp.line.width = Pt(0.75)
        na_shapes.append(new_shp)

    for i, (x, y) in enumerate(needed_slots):
        if i < len(na_shapes):
            shp = na_shapes[i]
            shp.left = x
            shp.top = y
            shp.width = 383540
            shp.height = 245110
            if shp.has_text_frame:
                shp.text_frame.text = "N/A"
                p = shp.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                if p.runs:
                    p.runs[0].font.size = Pt(8)
                    p.runs[0].font.bold = True
                    p.runs[0].font.name = "Arial"
                    p.runs[0].font.color.rgb = RGBColor(160, 160, 160)

    for j in range(len(needed_slots), len(na_shapes)):
        shp = na_shapes[j]
        shp.left = -2000000  # move offscreen

def update_bottom_trend_charts(prs, copq):
    """Update Charts 32, 34, 36, 38, 40 at the bottom of Slide 0."""
    slide0 = list(prs.slides)[0]

    def _format_bottom_labels(ch):
        # The template rotates category text by -1000 degrees and enables
        # ellipsis in a narrow margin. Reserve room and show horizontal site
        # names on both the quantity and working-hours charts.
        plot_area = ch._chartSpace.chart.plotArea
        layout = plot_area.find(qn('c:layout'))
        if layout is not None:
            manual = layout.find(qn('c:manualLayout'))
            if manual is not None:
                for tag, value in (('x', '0.25'), ('y', '0.22'), ('w', '0.70'), ('h', '0.63')):
                    element = manual.find(qn('c:' + tag))
                    if element is not None:
                        element.set('val', value)
        for axis in plot_area.iter(qn('c:catAx')):
            hidden = axis.find(qn('c:delete'))
            if hidden is not None:
                hidden.set('val', '0')
            for body in axis.iter(qn('a:bodyPr')):
                body.set('rot', '0')
                body.set('vert', 'horz')
                body.set('wrap', 'none')
                body.set('vertOverflow', 'overflow')
                for inset in ('lIns', 'rIns', 'tIns', 'bIns'):
                    body.set(inset, '0')
            for font in axis.iter(qn('a:defRPr')):
                font.set('sz', '650')
                font.set('b', '0')
            for latin in axis.iter(qn('a:latin')):
                latin.set('typeface', 'Arial')
        for ser in _ser_elements(ch):
            for dlbls in ser.findall(qn('c:dLbls')):
                pos = dlbls.find(qn('c:dLblPos'))
                if pos is None:
                    pos = etree.SubElement(dlbls, qn('c:dLblPos'))
                pos.set('val', 'outEnd')
                for dl in dlbls.findall(qn('c:dLbl')):
                    for ml in dl.findall(qn('c:layout')):
                        dl.remove(ml)
                    for ext in dl.findall(qn('c:extLst')):
                        dl.remove(ext)
                tx_prs = list(dlbls.iter(qn('c:txPr')))
                if not tx_prs:
                    tx_prs = [etree.SubElement(dlbls, qn('c:txPr'))]
                for tx_pr in tx_prs:
                    for body_pr in tx_pr.iter(qn('a:bodyPr')):
                        body_pr.set('wrap', 'none')
                        body_pr.set('lIns', '0')
                        body_pr.set('rIns', '0')
                        body_pr.set('tIns', '0')
                        body_pr.set('bIns', '0')
                        body_pr.set('vertOverflow', 'clip')
                    for latin in tx_pr.iter(qn('a:latin')):
                        latin.set('typeface', 'Arial')
                    for def_rpr in tx_pr.iter(qn('a:defRPr')):
                        def_rpr.set('sz', '650')
                        def_rpr.set('b', '1')
        strip_filtered_extensions(ch)

    # Chart 32: B/C Qty (prs) [JV2, JV, VH2, VH]
    ch32 = get_chart_by_name(slide0, "Chart 32")
    if ch32:
        vals = []
        for s in ["JV2", "JV", "VH2", "VH"]:
            df = copq.get(s)
            v = df[df.apply(map_bar_cat, axis=1) == "B/C Grade"]["Defective Qty(Pair)"].sum() if df is not None and not df.empty else 0.0
            vals.append(round(v, 1) if v > 0.001 else None)
        for ser in _ser_elements(ch32):
            set_series_values_ordered(ser, vals, format_code='#,##0.0')
        _format_bottom_labels(ch32)
        log.info(f"Chart 32 (B/C Qty): {vals}")

    # Chart 34: Re-inspection Qty (prs) [JVB, JV2, JV, VH3, VH2, VH]
    ch34 = get_chart_by_name(slide0, "Chart 34")
    if ch34:
        vals = []
        for s in ["JVB", "JV2", "JV", "VH3", "VH2", "VH"]:
            df = copq.get(s)
            v = df[df.apply(map_bar_cat, axis=1) == "Re-inspection"]["Defective Qty(Pair)"].sum() if df is not None and not df.empty else 0.0
            vals.append(round(v, 0) if v > 0.001 else None)
        for ser in _ser_elements(ch34):
            set_series_values_ordered(ser, vals, format_code='#,##0')
        _format_bottom_labels(ch34)
        _set_chart_valax_max(ch34, 185000)
        log.info(f"Chart 34 (Re-insp Qty): {vals}")

    # Chart 36: Re-inspection Work hours [JVB, JV2, JV, VH3, VH2, VH]
    ch36 = get_chart_by_name(slide0, "Chart 36")
    if ch36:
        vals = []
        for s in ["JVB", "JV2", "JV", "VH3", "VH2", "VH"]:
            df = copq.get(s)
            v = df[df.apply(map_bar_cat, axis=1) == "Re-inspection"]["Working hours"].sum() if df is not None and not df.empty else 0.0
            vals.append(round(v, 1) if v > 0.001 else None)
        for ser in _ser_elements(ch36):
            set_series_values_ordered(ser, vals, format_code='#,##0.0')
        _format_bottom_labels(ch36)
        log.info(f"Chart 36 (Re-insp Hours): {vals}")

    # Chart 38: Touch-up Qty (prs) [JV, VH4, VH2, VH]
    ch38 = get_chart_by_name(slide0, "Chart 38")
    if ch38:
        vals = []
        for s in ["JV", "VH4", "VH2", "VH"]:
            df = copq.get(s)
            v = df[df.apply(map_bar_cat, axis=1) == "Touch up paint"]["Defective Qty(Pair)"].sum() if df is not None and not df.empty else 0.0
            vals.append(round(v, 0) if v > 0.001 else None)
        for ser in _ser_elements(ch38):
            set_series_values_ordered(ser, vals, format_code='#,##0')
        _format_bottom_labels(ch38)
        _set_chart_valax_max(ch38, 185000)
        log.info(f"Chart 38 (Touch-up Qty): {vals}")

    # Chart 40: Touch-up Work hours [JV, VH4, VH2, VH]
    ch40 = get_chart_by_name(slide0, "Chart 40")
    if ch40:
        vals = []
        for s in ["JV", "VH4", "VH2", "VH"]:
            df = copq.get(s)
            v = df[df.apply(map_bar_cat, axis=1) == "Touch up paint"]["Working hours"].sum() if df is not None and not df.empty else 0.0
            vals.append(round(v, 1) if v > 0.001 else None)
        for ser in _ser_elements(ch40):
            set_series_values_ordered(ser, vals, format_code='#,##0.0')
        _format_bottom_labels(ch40)
        log.info(f"Chart 40 (Touch-up Hours): {vals}")
        for ser in _ser_elements(ch40):
            set_series_values_ordered(ser, vals, format_code='#,##0.0')
        _format_bottom_labels(ch40)
        log.info(f"Chart 40 (Touch-up Hours): {vals}")

def update_slide1_monthly(prs):
    """Update Slide 1 monthly trend charts (VH: Chart 1, VH2: Chart 6, JV: Chart 9, JV2: Chart 8)
    from Database/CoPQ_type_analysis.xlsx rolling 7-month history + current month."""
    if not os.path.exists(DB_TYPE):
        log.warning(f"DB_TYPE not found: {DB_TYPE}")
        return
    import openpyxl
    wb = openpyxl.load_workbook(DB_TYPE, data_only=True)
    if "Current month" not in wb.sheetnames:
        return
    ws = wb["Current month"]

    # Year-to-date months: from January (col 18 for 2026) to current month (e.g. Aug, col 25)
    curr_col = None
    for c in range(3, ws.max_column + 1):
        hdr = ws.cell(2, c).value
        cost = str(ws.cell(27, c).value or "")
        if hdr and not cost.startswith("="):
            curr_col = c
    if curr_col is None:
        curr_col = 25

    jan_col = None
    for c in range(curr_col, 2, -1):
        hdr = str(ws.cell(2, c).value or "").strip().lower()
        if hdr.startswith("jan"):
            jan_col = c
            break
    if jan_col is None:
        jan_col = 18

    month_cols = list(range(jan_col, curr_col + 1))
    if len(month_cols) > 7:
        month_cols = month_cols[-7:]
    months = [str(ws.cell(2, c).value or "") for c in month_cols]

    site_row_map = {
        "VH":  {"name": "Chart 1", "qty_row": 3, "cost_row": 10},
        "VH2": {"name": "Chart 6", "qty_row": 4, "cost_row": 11},
        "JV":  {"name": "Chart 9", "qty_row": 5, "cost_row": 12},
        "JV2": {"name": "Chart 8", "qty_row": 6, "cost_row": 13},
    }

    slide1 = list(prs.slides)[1]
    for site, info in site_row_map.items():
        ch = get_chart_by_name(slide1, info["name"])
        if ch is None:
            continue
        qty_vals = [float(ws.cell(info["qty_row"], c).value or 0) for c in month_cols]
        cost_vals = [float(ws.cell(info["cost_row"], c).value or 0) for c in month_cols]

        sers = _ser_elements(ch)
        if len(sers) >= 2:
            set_series_categories(sers[0], months)
            set_series_categories(sers[1], months)
            set_series_values_ordered(sers[0], qty_vals)
            set_series_values_ordered(sers[1], cost_vals, format_code='"$"#,##0')

            _format_monthly_qty_labels(sers[0], qty_vals=qty_vals, rot=0, font_sz=700)
            _format_monthly_cost_labels(sers[1], cost_vals=cost_vals, rot=0, font_sz=750)
            strip_filtered_extensions(ch)
            log.info(f"Slide 1 monthly trend {info['name']} ({site}): Qty={qty_vals[-1]}, Cost=${cost_vals[-1]:,.2f}")

def _format_monthly_qty_labels(ser, qty_vals=None, rot=0, font_sz=700):
    """Format data labels for monthly trend qty series horizontally with bottom position.
    Suppresses 0-value labels to avoid collision at zero axis."""
    dlbls = ser.xpath('.//c:dLbls')
    if dlbls:
        dl = dlbls[0]
        for ind_dl in dl.findall(qn('c:dLbl')):
            for ml in ind_dl.findall(qn('c:layout')):
                ind_dl.remove(ml)
            for ext in ind_dl.findall(qn('c:extLst')):
                ind_dl.remove(ext)
            pos_i = ind_dl.find(qn('c:dLblPos'))
            if pos_i is not None:
                pos_i.set('val', 'b')

        if qty_vals:
            for i, q in enumerate(qty_vals):
                if q == 0 or q < 0.001:
                    d_del = etree.Element(qn('c:dLbl'))
                    etree.SubElement(d_del, qn('c:idx'), val=str(i))
                    etree.SubElement(d_del, qn('c:delete'), val='1')
                    dl.insert(0, d_del)

        pos = dl.find(qn('c:dLblPos'))
        if pos is None:
            pos = etree.SubElement(dl, qn('c:dLblPos'))
        pos.set('val', 'b')
        txPr = dl.xpath('.//c:txPr')
        if txPr:
            bodyPr = txPr[0].xpath('.//a:bodyPr')
            if bodyPr:
                bodyPr[0].attrib['rot'] = str(rot)
                bodyPr[0].attrib['wrap'] = 'none'
                bodyPr[0].attrib['lIns'] = '0'
                bodyPr[0].attrib['rIns'] = '0'
            defRPr = txPr[0].xpath('.//a:defRPr')
            if defRPr:
                defRPr[0].attrib['sz'] = str(font_sz)
                defRPr[0].attrib['b'] = '1'

def _format_monthly_cost_labels(ser, cost_vals=None, rot=0, font_sz=750):
    """Format data labels for monthly trend cost series horizontally with top position.
    Lifts 0-value labels slightly to avoid touching the zero axis marker."""
    dlbls = ser.xpath('.//c:dLbls')
    if dlbls:
        dl = dlbls[0]
        for ind_dl in dl.findall(qn('c:dLbl')):
            for ml in ind_dl.findall(qn('c:layout')):
                ind_dl.remove(ml)
            for ext in ind_dl.findall(qn('c:extLst')):
                ind_dl.remove(ext)
            pos_i = ind_dl.find(qn('c:dLblPos'))
            if pos_i is not None:
                pos_i.set('val', 't')

        if cost_vals:
            for i, c in enumerate(cost_vals):
                if c == 0 or c < 0.01:
                    d_lift = etree.Element(qn('c:dLbl'))
                    etree.SubElement(d_lift, qn('c:idx'), val=str(i))
                    lay = etree.SubElement(d_lift, qn('c:layout'))
                    ml = etree.SubElement(lay, qn('c:manualLayout'))
                    etree.SubElement(ml, qn('c:y'), val='-0.04')
                    etree.SubElement(d_lift, qn('c:dLblPos'), val='t')
                    etree.SubElement(d_lift, qn('c:showLegendKey'), val='0')
                    etree.SubElement(d_lift, qn('c:showVal'), val='1')
                    etree.SubElement(d_lift, qn('c:showCatName'), val='0')
                    etree.SubElement(d_lift, qn('c:showSerName'), val='0')
                    etree.SubElement(d_lift, qn('c:showPercent'), val='0')
                    etree.SubElement(d_lift, qn('c:showBubbleSize'), val='0')
                    dl.insert(0, d_lift)

        pos = dl.find(qn('c:dLblPos'))
        if pos is None:
            pos = etree.SubElement(dl, qn('c:dLblPos'))
        pos.set('val', 't')
        txPr = dl.xpath('.//c:txPr')
        if txPr:
            bodyPr = txPr[0].xpath('.//a:bodyPr')
            if bodyPr:
                bodyPr[0].attrib['rot'] = str(rot)
                bodyPr[0].attrib['wrap'] = 'none'
                bodyPr[0].attrib['lIns'] = '0'
                bodyPr[0].attrib['rIns'] = '0'
            defRPr = txPr[0].xpath('.//a:defRPr')
            if defRPr:
                defRPr[0].attrib['sz'] = str(font_sz)
                defRPr[0].attrib['b'] = '1'


def update_pie_chart_data(ch, cat_val_pairs, default_colors=None):
    """
    Updates a pie chart with non-zero cat_val_pairs: [(cat, val), ...].
    Preserves existing dPt colors where matching, removes zero/unused categories so
    0-value slices and their $0.00 data labels never show up on the pie chart!
    """
    series = _ser_elements(ch)
    if not series:
        return
    ser = series[0]
    dpts = ser.findall(qn('c:dPt'))

    active_pairs = [(c, round(float(v), 2)) for c, v in cat_val_pairs if v is not None and float(v) > 0.0001]
    if not active_pairs:
        set_series_categories(ser, [""])
        set_series_values_ordered(ser, [None])
        for dp in dpts:
            ser.remove(dp)
        for dlbls in ser.findall(qn('c:dLbls')):
            for dl in list(dlbls):
                dlbls.remove(dl)
            dlbls.append(etree.Element(qn('c:delete'), val='1'))
        strip_filtered_extensions(ch)
        return

    new_cats = [c for c, _ in active_pairs]
    new_vals = [v for _, v in active_pairs]

    set_series_categories(ser, new_cats)
    set_series_values_ordered(ser, new_vals)

    for dp in dpts:
        ser.remove(dp)

    for i, cat in enumerate(new_cats):
        col_hex = PIE_COLORS[i % len(PIE_COLORS)]
        dp = etree.SubElement(ser, qn('c:dPt'))
        etree.SubElement(dp, qn('c:idx'), val=str(i))
        spPr = etree.SubElement(dp, qn('c:spPr'))
        solidFill = etree.SubElement(spPr, qn('a:solidFill'))
        etree.SubElement(solidFill, qn('a:srgbClr'), val=col_hex)
        if len(new_cats) > 1:
            line = etree.SubElement(spPr, qn('a:ln'), w='12700')
            line_fill = etree.SubElement(line, qn('a:solidFill'))
            etree.SubElement(line_fill, qn('a:srgbClr'), val='FFFFFF')

    # Remove old manual offsets and individual txPr overrides in dLbl so PowerPoint auto bestFit avoids wrapping
    for dlbls in ser.findall(qn('c:dLbls')):
        for dl in dlbls.findall(qn('c:dLbl')):
            for ml in dl.findall(qn('c:layout')):
                dl.remove(ml)
            for tx in dl.findall(qn('c:txPr')):
                dl.remove(tx)
        txPr = dlbls.find(qn('c:txPr'))
        if txPr is None:
            txPr = etree.SubElement(dlbls, qn('c:txPr'))
        bodyPr = txPr.find(qn('a:bodyPr'))
        if bodyPr is None:
            bodyPr = etree.SubElement(txPr, qn('a:bodyPr'))
        bodyPr.set('wrap', 'none')
        for defRPr in txPr.iter(qn('a:defRPr')):
            defRPr.set('sz', '750')
            defRPr.set('b', '1')

    strip_filtered_extensions(ch)

def update_factory_pie(prs, copq):
    slide1 = list(prs.slides)[1]
    # Safety layout: adjust JV2 pie chart height/top so it does not collide with footer
    ch_jv2 = get_chart_by_name(slide1, "Chart 33")
    if ch_jv2:
        shp_jv2 = next((s for s in slide1.shapes if s.name == "Chart 33"), None)
        if shp_jv2:
            shp_jv2.top = 5120000
            shp_jv2.height = 1300000

    for g in SLIDE1_GROUPS:
        site = g["site"]
        ch = get_chart_by_name(slide1, g["factory_pie"])
        if ch is None:
            WARN.append(f"factory pie '{g['factory_pie']}' ({site}) not found"); continue
        df = copq.get(site)
        if df is None or df.empty:
            update_pie_chart_data(ch, [])
            continue
        bc = df[df.apply(map_bar_cat, axis=1) == "B/C Grade"]
        if bc.empty or bc["Ttl cost ($)"].sum() < 0.01:
            update_pie_chart_data(ch, [])
            continue
        fac_sum = bc.groupby("Factory")["Ttl cost ($)"].sum().to_dict()
        cat_val_pairs = sorted([(str(k), float(v)) for k, v in fac_sum.items()], key=lambda x: x[1], reverse=True)
        update_pie_chart_data(ch, cat_val_pairs)
        log.info(f"factory pie ({site}) updated with active slices: {cat_val_pairs}")

def update_top3_model(prs, copq):
    top5_charts = {"VH": "Chart 24", "VH2": "Chart 35",
                   "JV": "Chart 38", "JV2": "Chart 40"}
    slide1 = list(prs.slides)[1]
    copq_top_models = {}
    for site, name in top5_charts.items():
        ch = get_chart_by_name(slide1, name)
        if ch is None:
            WARN.append(f"top3 model '{name}' ({site}) not found"); continue
        df = copq.get(site)
        if df is None or df.empty:
            for ser in _ser_elements(ch):
                set_series_categories(ser, ["", "", "", "", ""])
                set_series_values_ordered(ser, [None, None, None, None, None])
            strip_filtered_extensions(ch)
            WARN.append(f"top3 model ({site}): no COPQ data"); continue
        bc = df[df.apply(map_bar_cat, axis=1) == "B/C Grade"]
        if bc.empty or bc["Ttl cost ($)"].sum() < 0.01:
            for ser in _ser_elements(ch):
                set_series_categories(ser, ["", "", "", "", ""])
                set_series_values_ordered(ser, [None, None, None, None, None])
            strip_filtered_extensions(ch)
            continue
        n = 5
        top = bc.groupby("Model")["Ttl cost ($)"].sum().sort_values(ascending=False).head(n)
        copq_top_models[site] = top.head(3).index.tolist()
        top_asc = top.iloc[::-1]
        cats = [clean_model_name(str(m), max_len=20) for m in top_asc.index]
        vals = [round(float(v), 2) for v in top_asc.values]
        while len(cats) < n:
            cats.insert(0, "")
            vals.insert(0, None)
        for ser in _ser_elements(ch):
            set_series_categories(ser, cats)
            set_series_values_ordered(ser, vals)

        format_model_chart(ch, rot=0, font_sz=650)

        # Enforce Safe Zone: Dynamically scale valAx max so horizontal bars and labels
        # never encroach or collide with the adjacent 'Top 3 Defect' stacked chart!
        valid_vals = [v for v in vals if v is not None and v > 0]
        if valid_vals:
            max_cost = max(valid_vals)
            default_min = 3000 if site == "JV" else 1500
            dynamic_max = max(default_min, max_cost * 4.2)
            rounded_max = int(math.ceil(dynamic_max / 1000.0) * 1000)
            _set_chart_valax_max(ch, rounded_max)
            log.info(f"top model ({site}) -> valAx max scaled to {rounded_max} for safe zone")

        strip_filtered_extensions(ch)
        log.info(f"top model ({site}) -> {[c for c in cats if c]}")
    return copq_top_models


def _set_chart_valax_max(ch, max_val):
    """Set value axis maximum on chart XML to enforce a safe zone before adjacent charts."""
    xml = ch._element
    valAx = xml.xpath('.//c:valAx')
    if valAx:
        scaling = valAx[0].xpath('.//c:scaling')
        if scaling:
            max_elem = scaling[0].xpath('.//c:max')
            if max_elem:
                max_elem[0].attrib['val'] = str(int(max_val))
            else:
                etree.SubElement(scaling[0], qn('c:max'), val=str(int(max_val)))

def update_all_slide_titles(prs, period_label):
    """Update title and subtitles across all slides to reflect period_label."""
    # Slide 0 Title
    slide0 = prs.slides[0]
    for shp in slide0.shapes:
        if shp.has_text_frame and ("Cost of Poor Quality" in shp.text_frame.text or "CoPQ" in shp.text_frame.text):
            p = shp.text_frame.paragraphs[0]
            p.text = f"Cost of Poor Quality (CoPQ)_{period_label}"
            if p.runs:
                p.runs[0].font.name = "Arial"
                p.runs[0].font.size = Pt(20)
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = RGBColor(60, 77, 160)  # Chingluh Brand Navy #3C4DA0
            log.info(f"Slide 0 title -> 'Cost of Poor Quality (CoPQ)_{period_label}'")
            break

    # Slides 1, 2, 3 Subtitles
    for idx in (1, 2, 3):
        if idx < len(prs.slides):
            slide = prs.slides[idx]
            for shp in slide.shapes:
                if shp.has_text_frame:
                    txt = shp.text_frame.text.strip()
                    if txt == "CoPQ Analysis" or "CoPQ Analysis" in txt:
                        p = shp.text_frame.paragraphs[0]
                        p.text = "CoPQ Analysis"
                        break

def apply_chart_hyperlinks(prs, db_overview_rel, db_type_rel, ftt_rel=None):
    """
    Apply interactive hyperlinks to all charts and overview summary tables.
    Clicking any chart or table during viewing or presentation opens the
    corresponding live database Excel workbook:
      - Slide 0 (Overview, site bars, country pies, table) -> CoPQ database 25.xlsx
      - Slide 1 (B/C Grade details, trends, models, pies)   -> CoPQ_type_analysis.xlsx (or FTT for defects)
      - Slide 2 (Re-inspection models & pies)               -> CoPQ_type_analysis.xlsx
      - Slide 3 (Touch-up models & pies)                    -> CoPQ_type_analysis.xlsx
    """
    total_linked = 0
    slides = list(prs.slides)
    for s_idx, slide in enumerate(slides):
        for shp in slide.shapes:
            target_link = None
            if s_idx == 0:
                if shp.has_chart or shp.has_table or shp.name in ["Table 102", "Rectangle 68", "Rectangle 69", "Rectangle 73"]:
                    target_link = db_overview_rel
            elif s_idx == 1:
                if shp.has_chart:
                    if shp.name in ["Chart 78", "Chart 80", "Chart 92", "Chart 96"]:
                        target_link = ftt_rel if ftt_rel else db_type_rel
                    else:
                        target_link = db_type_rel
            elif s_idx in [2, 3]:
                if shp.has_chart:
                    target_link = db_type_rel

            if target_link:
                try:
                    shp.click_action.hyperlink.address = target_link
                    total_linked += 1
                except Exception:
                    pass
    log.info(f"Applied interactive hyperlinks to {total_linked} charts/tables.")
    return total_linked


def update_pptx_chart_data_sources(pptx_path, db_overview_path, db_type_path, ftt_report_path=None):
    """
    Update internal OLE/externalData relationship targets for all 51 charts inside pptx_path,
    pointing them to the local Database workbooks using file:/// URI format.
    In PowerPoint on Windows, 'Edit Data' resolves external links via full file:/// URIs
    so clicking 'Edit Data' opens the exact Excel workbook directly!
    """
    import zipfile
    import tempfile
    import xml.etree.ElementTree as ET

    db_overview_abs = os.path.abspath(db_overview_path).replace('/', '\\')
    db_type_abs = os.path.abspath(db_type_path).replace('/', '\\')
    ftt_abs = os.path.abspath(ftt_report_path or "Output/FTT_Combined_Report.xlsx").replace('/', '\\')

    db_overview_target = f"file:///{db_overview_abs}"
    db_type_target = f"file:///{db_type_abs}"
    ftt_target = f"file:///{ftt_abs}"

    temp_zip = pptx_path + ".tmp.zip"
    shutil.copy2(pptx_path, temp_zip)
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(temp_zip, 'r') as zin:
            zin.extractall(temp_dir)

        charts_rel_dir = os.path.join(temp_dir, 'ppt', 'charts', '_rels')
        if os.path.exists(charts_rel_dir):
            for rel_name in os.listdir(charts_rel_dir):
                if not rel_name.endswith('.xml.rels'):
                    continue
                rel_file = os.path.join(charts_rel_dir, rel_name)
                tree = ET.parse(rel_file)
                root = tree.getroot()
                changed = False
                for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    t = rel.attrib.get('Type', '')
                    if 'oleObject' in t or 'package' in t:
                        curr_target = rel.attrib.get('Target', '')
                        fn = os.path.basename(curr_target.replace("\\", "/")).lower()
                        if 'copq database' in fn or 'copq%20database' in fn or 'database%2025' in fn or 'database 25' in fn:
                            new_target = db_overview_target
                        elif 'type_analysis' in fn or 'type%20analysis' in fn:
                            new_target = db_type_target
                        elif 'ftt' in fn:
                            new_target = ftt_target
                        else:
                            new_target = db_type_target

                        rel.attrib['Target'] = new_target
                        rel.attrib['TargetMode'] = 'External'
                        changed = True
                if changed:
                    tree.write(rel_file, xml_declaration=True, encoding='utf-8')

        with zipfile.ZipFile(pptx_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    full_p = os.path.join(root_dir, file)
                    arc_p = os.path.relpath(full_p, temp_dir)
                    zout.write(full_p, arc_p)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if os.path.exists(temp_zip):
            os.remove(temp_zip)


def update_site_total_table(prs, copq):
    """Update Slide 0 Table 102 headers: Site Name on Line 1, Total on Line 2.
    Clean typography, bold, centered, no diff arrows."""
    slide0 = list(prs.slides)[0]
    site_names = ["VH", "VH2", "VH3", "VH4", "JV", "JV2", "JVB"]
    for shp in slide0.shapes:
        if not shp.has_table:
            continue
        tbl = shp.table
        for ci, site in enumerate(site_names, start=1):
            if ci >= len(tbl.rows[0].cells):
                break
            cell = tbl.rows[0].cells[ci]
            total = float(non_rework(copq[site])["Ttl cost ($)"].sum()) if site in copq and not copq[site].empty else 0.0
            tf = cell.text_frame
            tf.word_wrap = True
            p0 = tf.paragraphs[0]
            p0.text = f"{site} "
            p0.alignment = PP_ALIGN.CENTER
            if p0.runs:
                p0.runs[0].font.name = "Arial"
                p0.runs[0].font.size = Pt(11)
                p0.runs[0].font.bold = True
                p0.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            if len(tf.paragraphs) > 1:
                p1 = tf.paragraphs[1]
            else:
                p1 = tf.add_paragraph()
            p1.text = f"${total:,.2f}"
            p1.alignment = PP_ALIGN.CENTER
            if p1.runs:
                p1.runs[0].font.name = "Arial"
                p1.runs[0].font.size = Pt(11)
                p1.runs[0].font.bold = True
                p1.runs[0].font.color.rgb = RGBColor(0, 0, 0)
            log.info(f"table header {site}: ${total:,.2f}")
        return

def update_country_totals(prs, copq):
    vn_sites = [s for s in copq if s.startswith("VH")]
    id_sites = [s for s in copq if s.startswith("JV")]
    vn = sum(float(non_rework(copq[s])["Ttl cost ($)"].sum()) for s in vn_sites if s in copq and not copq[s].empty)
    idt = sum(float(non_rework(copq[s])["Ttl cost ($)"].sum()) for s in id_sites if s in copq and not copq[s].empty)
    clg = vn + idt
    slide0 = list(prs.slides)[0]
    total_boxes = {
        70: (vn, RGBColor(0, 0, 0), Pt(12.5)),         # VN: dark text on white circle background
        87: (clg, RGBColor(0, 0, 0), Pt(13)),          # CLG: black text
        89: (idt, RGBColor(0, 0, 0), Pt(12.5)),        # ID: dark text on white circle background
    }
    for shp in slide0.shapes:
        if shp.shape_id in total_boxes:
            val, color, sz = total_boxes[shp.shape_id]
            new = f"${val:,.2f}"
            if shp.has_text_frame and shp.text_frame.paragraphs:
                p = shp.text_frame.paragraphs[0]
                p.text = new
                p.alignment = PP_ALIGN.CENTER
                if p.runs:
                    p.runs[0].font.name = "Arial"
                    p.runs[0].font.size = sz
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = color
                log.info(f"total box id={shp.shape_id} -> {new}")
    for nm, val in (("Chart 57", vn), ("Chart 59", idt)):
        ch = get_chart_by_name(slide0, nm)
        if ch:
            for ser in _ser_elements(ch):
                set_series_values_by_category(ser, {"0": val}, warn=False)
            log.info(f"{nm} doughnut -> ${val:,.2f}")
    ch = get_chart_by_name(slide0, "Chart 58")
    if ch:
        for ser in _ser_elements(ch):
            set_series_values_by_category(ser, {"VN": vn, "Indo": idt}, warn=False)
        log.info(f"CLG doughnut -> VN ${vn:,.2f}, Indo ${idt:,.2f}")

def _load_prior_month_overview():
    """Load prior month overview figures from CoPQ database 25.xlsx for MoM delta calculations."""
    try:
        xl = pd.ExcelFile(DB_OVERVIEW)
        target_sheet = None
        try:
            import datetime as _dt
            curr_dt = _dt.datetime.strptime(DB_MONTH, "%b-%Y")
            prior_dt = curr_dt.replace(day=1) - _dt.timedelta(days=1)
            cands = [
                prior_dt.strftime("%b-%Y"),
                prior_dt.strftime("%B-%Y"),
                prior_dt.strftime("%b-%y"),
            ]
            for c in cands:
                if c in xl.sheet_names:
                    target_sheet = c
                    break
        except Exception:
            pass

        if not target_sheet:
            candidates = [s for s in xl.sheet_names if any(y in s for y in ["2025", "2026"]) and not s.startswith("COPQ_rows") and s != "Current" and s != DB_MONTH]
            if candidates:
                target_sheet = candidates[-1]

        if target_sheet:
            log.info(f"Loaded prior month overview from sheet: {target_sheet}")
            df = xl.parse(target_sheet, header=1).set_index("Factory")
            return df
    except Exception as e:
        log.warning(f"Could not load prior month overview: {e}")
    return None

def load_mom_diffs_from_analysis(analysis_path):
    """
    Extract exact evaluated MoM difference values from CoPQ_type_analysis.xlsx:
      - Re-inspection diffs (Rows 27-32)
      - Touch-up diffs (Rows 44-48)
    Formula in Difference column: curr_col - prev_col.
    """
    import openpyxl
    if not os.path.exists(analysis_path):
        return {}, {}
    try:
        wb = openpyxl.load_workbook(analysis_path, data_only=False)
        if 'Current month' not in wb.sheetnames:
            return {}, {}
        ws = wb['Current month']

        col_diff = None
        for c in range(3, ws.max_column + 1):
            val = str(ws.cell(27, c).value or '')
            if val.startswith('='):
                col_diff = c
                break

        if not col_diff:
            col_diff = 26

        prev_col = col_diff - 2
        curr_col = col_diff - 1

        re_diffs = {}
        for r, s in zip(range(27, 33), ["VH", "VH2", "VH3", "JV", "JV2", "JVB"]):
            c_val = float(ws.cell(r, curr_col).value or 0.0)
            p_val = float(ws.cell(r, prev_col).value or 0.0)
            re_diffs[s] = c_val - p_val

        tu_diffs = {}
        for r, s in zip(range(44, 49), ["VH", "VH2", "VH4", "JV", "JVB"]):
            c_val = float(ws.cell(r, curr_col).value or 0.0)
            p_val = float(ws.cell(r, prev_col).value or 0.0)
            tu_diffs[s] = c_val - p_val

        return re_diffs, tu_diffs
    except Exception as e:
        log.warning(f"Could not load MoM diffs from analysis: {e}")
        return {}, {}


def apply_delta_badge_styling(grp, diff_val, font_size_pt=8.5):
    """
    Update delta badge group (Rectangle + Triangle):
    - diff_val > 0: INCREASE (giá tăng) -> RED/Magenta (#E6007E), Up triangle (mũi tên đi lên: rot=10800000, flipV=1)
    - diff_val < 0: DECREASE (giá giảm) -> GREEN (#00B050), Down triangle (mũi tên đi xuống: rot=10800000, no flipV)
    - diff_val == 0: GREEN $0.00
    Text format: f"${abs(diff_val):,.2f}"
    """
    rect = None
    tri = None
    for shp in grp.shapes:
        if 'Rectangle' in shp.name:
            rect = shp
        elif 'Triangle' in shp.name:
            tri = shp

    if rect is None:
        return

    abs_str = f"${abs(diff_val):,.2f}"
    is_inc = (diff_val > 0)

    # Red/Magenta for increase, Green for decrease
    tri_color_hex = "E6007E" if is_inc else "00B050"
    txt_color_rgb = RGBColor(230, 0, 126) if is_inc else RGBColor(0, 176, 80)

    # 1. Update text & font
    rect.text_frame.text = abs_str
    p = rect.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    if p.runs:
        r = p.runs[0]
        r.font.name = "Arial"
        r.font.size = Pt(font_size_pt)
        r.font.bold = True
        r.font.color.rgb = txt_color_rgb

    # 2. Update triangle shape direction & color
    # In the template coordinate system, rot="10800000" with flipV="1" points UP (▲),
    # while rot="10800000" without flipV points DOWN (▼).
    if tri is not None:
        xfrm = tri._element.find('.//' + qn('a:xfrm'))
        if xfrm is not None:
            xfrm.attrib['rot'] = '10800000'
            if is_inc:
                xfrm.attrib['flipV'] = '1'
            else:
                if 'flipV' in xfrm.attrib:
                    del xfrm.attrib['flipV']

        spPr = tri._element.find(qn('p:spPr'))
        if spPr is not None:
            solidFill = spPr.find(qn('a:solidFill'))
            if solidFill is None:
                solidFill = etree.SubElement(spPr, qn('a:solidFill'))
            srgb = solidFill.find(qn('a:srgbClr'))
            if srgb is None:
                srgb = etree.SubElement(solidFill, qn('a:srgbClr'))
            srgb.attrib['val'] = tri_color_hex


def update_slide2_reinspection(prs, copq):
    """Update Slide 2: Re-Inspection analysis (6 site badges, 6 top-3 model charts, 6 pie charts)."""
    slide2 = list(prs.slides)[2]
    re_diffs, _ = load_mom_diffs_from_analysis(DB_TYPE)

    # 1. Site KPI badges & MoM delta badges
    badge_groups = {
        "Group 41": {"site": "VH",  "delta_grp": "Group 73"},
        "Group 45": {"site": "VH2", "delta_grp": "Group 5"},
        "Group 49": {"site": "VH3", "delta_grp": "Group 79"},
        "Group 63": {"site": "JV",  "delta_grp": "Group 8"},
        "Group 53": {"site": "JV2", "delta_grp": "Group 92"},
        "Group 57": {"site": "JVB", "delta_grp": "Group 67"},
    }
    for grp_name, info in badge_groups.items():
        site = info["site"]
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Re-inspection"] if df is not None and not df.empty else None
        tot = float(sub["Ttl cost ($)"].sum()) if sub is not None and not sub.empty else 0.0

        # Update main cost badge (Large Blue Font: Arial 24pt, Bold, Color #0000CC)
        for shp in slide2.shapes:
            if shp.name == grp_name:
                for sub_shp in shp.shapes:
                    if sub_shp.has_text_frame and "$" in sub_shp.text_frame.text:
                        sub_shp.text_frame.text = f"${tot:,.2f}"
                        p = sub_shp.text_frame.paragraphs[0]
                        p.alignment = PP_ALIGN.CENTER
                        if p.runs:
                            p.runs[0].font.name = "Arial"
                            p.runs[0].font.size = Pt(24)
                            p.runs[0].font.bold = True
                            p.runs[0].font.color.rgb = RGBColor(0, 0, 204) # #0000CC Blue
                log.info(f"Slide 2 badge {shp.name} ({site}): ${tot:,.2f}")

        # Update MoM delta badge from CoPQ_type_analysis.xlsx formula
        diff = re_diffs.get(site, 0.0)
        for shp in slide2.shapes:
            if shp.name == info["delta_grp"]:
                apply_delta_badge_styling(shp, diff, font_size_pt=8.5)
                log.info(f"Slide 2 delta {info['delta_grp']} ({site}): diff={diff:+,.2f}")

    # 2. Top-3 Models Column Charts
    model_charts = {
        "Chart 23": "VH",
        "Chart 31": "VH2",
        "Chart 34": "VH3",
        "Chart 36": "JV",
        "Chart 17": "JV2",
        "Chart 27": "JVB",
    }
    for name, site in model_charts.items():
        ch = get_chart_by_name(slide2, name)
        if ch is None:
            continue
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Re-inspection"] if df is not None and not df.empty else None
        if sub is None or sub.empty or sub["Ttl cost ($)"].sum() < 0.01:
            for ser in _ser_elements(ch):
                set_series_categories(ser, ["", "", ""])
                set_series_values_ordered(ser, [None, None, None])
            strip_filtered_extensions(ch)
            continue
        top = sub.groupby("Model")["Ttl cost ($)"].sum().sort_values(ascending=False).head(3)
        cats = [clean_model_name(str(m), max_len=22) for m in top.index]
        vals = [round(float(v), 2) for v in top.values]
        if len(cats) > 1:
            while len(cats) < 3:
                cats.append("")
                vals.append(None)
        for ser in _ser_elements(ch):
            set_series_categories(ser, cats)
            set_series_values_ordered(ser, vals)
            set_series_color(ser, "33CCFF")
        format_model_chart(ch, rot=-35, font_sz=650)
        valid_vals = [v for v in vals if v is not None and v > 0]
        if valid_vals:
            max_v = max(valid_vals)
            target_max = int(math.ceil((max_v * 1.25) / 100.0) * 100)
            if target_max < 200:
                target_max = int(math.ceil((max_v * 1.25) / 20.0) * 20)
            _set_chart_valax_max(ch, max(target_max, 50))
        strip_filtered_extensions(ch)
        log.info(f"Slide 2 models {name} ({site}): {cats} -> {vals}")

    # 3. Factory / Shopfloor Pie Charts
    pie_charts = {
        "Chart 22": "VH",
        "Chart 32": "VH2",
        "Chart 33": "VH3",
        "Chart 37": "JV",
        "Chart 19": "JV2",
        "Chart 21": "JVB",
    }
    for name, site in pie_charts.items():
        ch = get_chart_by_name(slide2, name)
        if ch is None:
            continue
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Re-inspection"] if df is not None and not df.empty else None
        if sub is None or sub.empty or sub["Ttl cost ($)"].sum() < 0.01:
            update_pie_chart_data(ch, [])
            continue
        col_to_group = "Factory" if "Factory" in sub.columns else "Fty/Line"
        fac_sum = sub.groupby(col_to_group)["Ttl cost ($)"].sum().to_dict()
        cat_val_pairs = sorted([(str(k), float(v)) for k, v in fac_sum.items()], key=lambda x: x[1], reverse=True)[:5]
        update_pie_chart_data(ch, cat_val_pairs)
        log.info(f"Slide 2 pie {name} ({site}): {cat_val_pairs}")

def update_slide3_touchup(prs, copq):
    """Update Slide 3: Touch-Up Paint analysis (4 site badges, 4 top-3 model charts, 4 pie charts)."""
    slide3 = list(prs.slides)[3]
    _, tu_diffs = load_mom_diffs_from_analysis(DB_TYPE)

    # 1. Site KPI badges & MoM delta badges
    badge_shapes = {
        "Rectangle 23":  {"site": "VH",  "delta_grp": "Group 1"},
        "Rectangle 110": {"site": "VH2", "delta_grp": "Group 5"},
        "Rectangle 60":  {"site": "VH4", "delta_grp": "Group 120"},
        "Rectangle 111": {"site": "JV",  "delta_grp": "Group 41"},
    }
    for shp_name, info in badge_shapes.items():
        site = info["site"]
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Touch up paint"] if df is not None and not df.empty else None
        tot = float(sub["Ttl cost ($)"].sum()) if sub is not None and not sub.empty else 0.0

        # Update main cost badge (Large Blue Font: Arial 24pt, Bold, Color #0000CC)
        for shp in slide3.shapes:
            if shp.name == shp_name and shp.has_text_frame:
                shp.text_frame.text = f"${tot:,.2f}"
                p = shp.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                if p.runs:
                    p.runs[0].font.name = "Arial"
                    p.runs[0].font.size = Pt(24)
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = RGBColor(0, 0, 204) # #0000CC Blue
                log.info(f"Slide 3 badge {shp.name} ({site}): ${tot:,.2f}")

        # Update MoM delta badge from CoPQ_type_analysis.xlsx formula
        diff = tu_diffs.get(site, 0.0)
        for shp in slide3.shapes:
            if shp.name == info["delta_grp"]:
                apply_delta_badge_styling(shp, diff, font_size_pt=8.5)
                log.info(f"Slide 3 delta {info['delta_grp']} ({site}): diff={diff:+,.2f}")

    # 2. Top Models Column Charts
    model_charts = {
        "Chart 4":  "VH",
        "Chart 13": "VH2",
        "Chart 15": "VH4",
        "Chart 17": "JV",
    }
    for name, site in model_charts.items():
        ch = get_chart_by_name(slide3, name)
        if ch is None:
            continue
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Touch up paint"] if df is not None and not df.empty else None
        if sub is None or sub.empty or sub["Ttl cost ($)"].sum() < 0.01:
            for ser in _ser_elements(ch):
                set_series_categories(ser, ["", "", ""])
                set_series_values_ordered(ser, [None, None, None])
            strip_filtered_extensions(ch)
            continue
        top = sub.groupby("Model")["Ttl cost ($)"].sum().sort_values(ascending=False).head(3)
        cats = [clean_model_name(str(m), max_len=22) for m in top.index]
        vals = [round(float(v), 2) for v in top.values]
        if len(cats) > 1:
            while len(cats) < 3:
                cats.append("")
                vals.append(None)
        for ser in _ser_elements(ch):
            set_series_categories(ser, cats)
            set_series_values_ordered(ser, vals)
            set_series_color(ser, "33CCFF")
        format_model_chart(ch, rot=-35, font_sz=650)
        valid_vals = [v for v in vals if v is not None and v > 0]
        if valid_vals:
            max_v = max(valid_vals)
            target_max = int(math.ceil((max_v * 1.25) / 100.0) * 100)
            if target_max < 200:
                target_max = int(math.ceil((max_v * 1.25) / 20.0) * 20)
            _set_chart_valax_max(ch, max(target_max, 50))
        strip_filtered_extensions(ch)
        log.info(f"Slide 3 models {name} ({site}): {cats} -> {vals}")

    # 3. Shopfloor / Factory Pie Charts
    pie_charts = {
        "Chart 9":  "VH",
        "Chart 11": "VH2",
        "Chart 19": "VH4",
        "Chart 21": "JV",
    }
    for name, site in pie_charts.items():
        ch = get_chart_by_name(slide3, name)
        if ch is None:
            continue
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Touch up paint"] if df is not None and not df.empty else None
        if sub is None or sub.empty or sub["Ttl cost ($)"].sum() < 0.01:
            update_pie_chart_data(ch, [])
            continue
        col_to_group = "Factory" if "Factory" in sub.columns else "Fty/Line"
        fac_sum = sub.groupby(col_to_group)["Ttl cost ($)"].sum().to_dict()
        cat_val_pairs = sorted([(str(k), float(v)) for k, v in fac_sum.items()], key=lambda x: x[1], reverse=True)[:5]
        update_pie_chart_data(ch, cat_val_pairs)
        log.info(f"Slide 3 pie {name} ({site}): {cat_val_pairs}")

# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def clone(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)
    log.info(f"cloned {src}\n      -> {dst}")

def verify_opens(path):
    p = Presentation(path)
    log.info(f"VERIFY: re-opens OK, {len(p.slides)} slides")
    return len(p.slides)

def load_copq():
    xl = pd.ExcelFile(COPQ_XLSX)
    out = {}

    # Prefer reading directly from the unified master sheet 'COPQ_Clean' if available
    if "COPQ_Clean" in xl.sheet_names:
        df_master = xl.parse("COPQ_Clean", header=0)
        if "Ttl cost ($)" not in df_master.columns and "Cost" in df_master.columns:
            df_master["Ttl cost ($)"] = pd.to_numeric(df_master["Cost"], errors="coerce").fillna(0.0)
        if "Defective Qty(Pair)" not in df_master.columns and "Qty" in df_master.columns:
            df_master["Defective Qty(Pair)"] = pd.to_numeric(df_master["Qty"], errors="coerce").fillna(0.0)
        if "Working hours" not in df_master.columns:
            df_master["Working hours"] = 0.0

        for site, sub in df_master.groupby("Site"):
            sub = sub.copy()
            if "Date" in sub.columns:
                sub = sub[sub["Date"].notna()]
            if EXCLUDE_TYPES:
                mask = ~sub["Type"].astype(str).str.strip().isin(EXCLUDE_TYPES)
                sub = sub[mask]
            for c in ("Ttl cost ($)", "Defective Qty(Pair)", "Working hours"):
                if c in sub.columns:
                    sub[c] = pd.to_numeric(sub[c], errors="coerce").fillna(0.0)
            deck_site = DATA_SITE_ALIAS.get(str(site).strip(), str(site).strip())
            out[deck_site] = sub
        log.info(f"COPQ loaded from COPQ_Clean master: {len(out)} sites")
        return out

    # Fallback to individual per-site sheets
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=1)
        if "Date" not in df.columns and len(df) and str(df.iloc[0, 0]).startswith("Cost Of Poor Quality"):
            df = xl.parse(sheet, header=2)
        if "Date" in df.columns:
            df = df[df["Date"].notna()]
        if EXCLUDE_TYPES:
            mask = ~df["Type"].astype(str).str.strip().isin(EXCLUDE_TYPES)
            df = df[mask]
        if "Ttl cost ($)" not in df.columns and "Cost" in df.columns:
            df["Ttl cost ($)"] = pd.to_numeric(df["Cost"], errors="coerce").fillna(0.0)
        if "Defective Qty(Pair)" not in df.columns and "Qty" in df.columns:
            df["Defective Qty(Pair)"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0.0)
        if "Working hours" not in df.columns:
            df["Working hours"] = 0.0
        for c in ("Ttl cost ($)", "Defective Qty(Pair)", "Working hours"):
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
        deck_site = DATA_SITE_ALIAS.get(sheet, sheet)
        out[deck_site] = df
    return out

def load_copq_db():
    try:
        xl = pd.ExcelFile(DB_OVERVIEW)
    except FileNotFoundError:
        log.error(f"DB overview not found: {DB_OVERVIEW}")
        return {}
    # Use COPQ_Clean.xlsx master data directly
    return load_copq()

def generate_copq_presentation(
    period_label=None,
    db_month=None,
    out_pptx=None,
    copq_xlsx=None,
    ftt_xlsx=None,
    src_pptx=None,
    db_overview=None,
    db_type=None,
    allow_linked_ftt=True
):
    global PERIOD_LABEL, DB_MONTH, OUT_PPTX, COPQ_XLSX, FTT_XLSX, SRC_PPTX, DB_OVERVIEW, DB_TYPE

    if period_label:
        PERIOD_LABEL = period_label
    if db_month:
        DB_MONTH = db_month
    if out_pptx:
        OUT_PPTX = out_pptx
    if copq_xlsx:
        COPQ_XLSX = copq_xlsx
    if ftt_xlsx:
        FTT_XLSX = ftt_xlsx
    if src_pptx:
        SRC_PPTX = src_pptx
    if db_overview:
        DB_OVERVIEW = db_overview
    if db_type:
        DB_TYPE = db_type

    log.info("=" * 70)
    log.info(f"COPQ PPTX Executive Generator & Visualizer (Period: {PERIOD_LABEL})")
    log.info("=" * 70)
    if os.path.abspath(OUT_PPTX) == os.path.abspath(SRC_PPTX):
        raise SystemExit("Refusing to overwrite the source file.")

    target = OUT_PPTX
    if os.path.exists(target):
        try:
            os.remove(target)
        except PermissionError:
            import datetime as _dt
            stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            base, ext = os.path.splitext(OUT_PPTX)
            target = f"{base}_{stamp}{ext}"
            log.warning(f"Target is locked (open in another app). "
                        f"Writing to: {target}")

    clone(SRC_PPTX, target)
    copq = load_copq_db()
    if not copq:
        log.warning("No COPQ data loaded from database; falling back to COPQ_Clean.xlsx")
        copq = load_copq()
    prs = Presentation(target)

    # 1. Update Titles & Subtitles across all 4 slides
    update_all_slide_titles(prs, PERIOD_LABEL)

    # 2. Slide 0: Executive Overview
    update_site_total_table(prs, copq)
    update_country_totals(prs, copq)
    update_site_bar(prs, copq)
    update_na_badges(prs, copq)
    update_bottom_trend_charts(prs, copq)

    # 3. Slide 1: B/C Grade Detail
    update_slide1_monthly(prs)
    update_factory_pie(prs, copq)
    update_top3_model(prs, copq)
    # FTT charts and legends are transferred from Excel by Office after all
    # python-pptx edits. Their contents and formatting are never regenerated.

    # 4. Slide 2: Re-Inspection Detail
    update_slide2_reinspection(prs, copq)

    # 5. Slide 3: Touch-Up Paint Detail
    update_slide3_touchup(prs, copq)

    # 6. Apply reference-deck geometry; native FTT transfer follows below.
    apply_analysis_layout(prs)

    # 7. Apply live interactive hyperlinks to all charts & summary tables
    db_rel_overview = os.path.relpath(DB_OVERVIEW, os.path.dirname(target))
    db_rel_type = os.path.relpath(DB_TYPE, os.path.dirname(target))
    ftt_rel = os.path.relpath(FTT_XLSX, os.path.dirname(target)) if os.path.exists(FTT_XLSX) else None
    apply_chart_hyperlinks(prs, db_rel_overview, db_rel_type, ftt_rel)

    prs.save(target)

    # 8. Update internal chart OLE relationship targets (pass full absolute paths DB_OVERVIEW and DB_TYPE)
    update_pptx_chart_data_sources(target, DB_OVERVIEW, DB_TYPE, FTT_XLSX)
    transfer = copy_ftt_charts(target, FTT_XLSX, allow_linked=allow_linked_ftt)
    target = transfer['output']
    log.info(f"FTT charts copied from Excel: {transfer}")
    log.info("-" * 70)
    verify_opens(target)
    log.info("=" * 70)
    log.info(f"OUTPUT WRITTEN: {target}")
    log.info("=" * 70)
    log.info(f"WARNINGS: {len(WARN)}")
    for w in WARN:
        log.warning("  - " + w)
    log.info("=" * 70)
    return target

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Update COPQ presentation deck.")
    parser.add_argument("--period", default=None, help="Period label (e.g. 'Aug, 2026')")
    parser.add_argument("--db-month", default=None, help="DB month sheet name (e.g. 'Aug-2026')")
    parser.add_argument("--output", default=None, help="Target PPTX path")
    parser.add_argument("--copq-xlsx", default=None, help="COPQ_Clean.xlsx path")
    parser.add_argument("--ftt-xlsx", default=None, help="FTT_Combined_Report.xlsx path")
    parser.add_argument("--template", default=None, help="Template PPTX path")
    args = parser.parse_args()

    generate_copq_presentation(
        period_label=args.period,
        db_month=args.db_month,
        out_pptx=args.output,
        copq_xlsx=args.copq_xlsx,
        ftt_xlsx=args.ftt_xlsx,
        src_pptx=args.template
    )

if __name__ == "__main__":
    main()
