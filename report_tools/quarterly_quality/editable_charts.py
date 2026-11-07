"""Transfer Excel charts into QBR picture slots without rebuilding the deck."""
import ast
from io import BytesIO
from pathlib import Path
import math
import re
import posixpath
import textwrap
from zipfile import ZipFile

from lxml import etree as E
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter

C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PK = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = dict(c=C, a=A, p=P)
PERCENT_FORMAT = '0.0%'
RANK_PLOT_X = .04
RANK_PLOT_WIDTH = .92
# Excel source identity is based on source rows, never OOXML part numbers.
# Office renumbers chart parts when a template gains or loses a chart.
CHART_SPECS = (
    ('FTT', 'VN', 'trend', 2, 'FTT month-quarter', (36, 31, 32)),
    ('FTT', 'Indo', 'trend', 5, 'FTT month-quarter', (37, 33, 34)),
    ('FTT', 'VN', 'model', 6, 'FTT month-quarter', (41,)),
    ('FTT', 'Indo', 'model', 3, 'FTT month-quarter', (38,)),
    ('FTT', 'VN', 'defect', 7, 'FTT month-quarter', (51,)),
    ('FTT', 'Indo', 'defect', 13, 'FTT month-quarter', (48,)),
    ('HFPA', 'VN', 'trend', 2, 'Month HFPA', (28, 23, 24)),
    ('HFPA', 'Indo', 'trend', 3, 'Month HFPA', (29, 25, 26)),
    ('HFPA', 'VN', 'model', 8, 'Quarter HFPA ', (7,)),
    ('HFPA', 'Indo', 'model', 11, 'Quarter HFPA ', (4,)),
    ('HFPA', 'VN', 'defect', 10, 'Quarter HFPA ', (17,)),
    ('HFPA', 'Indo', 'defect', 12, 'Quarter HFPA ', (14,)),
)


def excel_chart_sources(book):
    sources = {}
    for name in book.namelist():
        if not re.fullmatch(r'xl/charts/chart\d+\.xml', name):
            continue
        root = E.fromstring(book.read(name))
        refs = root.xpath('.//c:ser/c:val/c:numRef/c:f/text()', namespaces=NS)
        try:
            parsed = [formula.rsplit('!', 1) for formula in refs]
            sheets = {owner.strip("'").replace("''", "'") for owner, _ in parsed}
            rows = tuple(range_boundaries(address)[1] for _, address in parsed)
        except (ValueError, TypeError):
            continue
        for metric, region, kind, _, sheet, expected_rows in CHART_SPECS:
            if sheets == {sheet} and rows == expected_rows:
                key = metric, region, kind
                if key in sources:
                    raise ValueError(f'Duplicate Excel chart for {key}.')
                sources[key] = name
    return sources


def compatible_trend_legend(data):
    """Keep the actual Excel legend in a separate band in WPS and PowerPoint."""
    root = E.fromstring(data)
    chart = root.find('c:chart', NS)
    plot = chart.find('c:plotArea', NS) if chart is not None else None
    if plot is None or plot.find('c:lineChart', NS) is None:
        return data
    for tx in plot.findall('.//c:ser/c:tx', NS):
        ref = tx.find('c:strRef', NS)
        formula = ref.find('c:f', NS) if ref is not None else None
        # Excel/PowerPoint accept a string formula here, but WPS does not
        # reliably render it as a series name. True cell references stay intact.
        if formula is not None and re.fullmatch(r'"(?:[^"]|"")*"', formula.text or ''):
            label = formula.text[1:-1].replace('""', '"')
            tx.remove(ref)
            node(tx, 'v').text = label
    legend = chart.find('c:legend', NS)
    if legend is None:
        legend = E.Element('{'+C+'}legend')
        chart.insert(chart.index(plot) + 1, legend)
    for child in list(legend):
        if E.QName(child).localname in ('legendPos', 'layout', 'overlay', 'txPr'):
            legend.remove(child)
    legend.insert(0, E.Element('{'+C+'}legendPos', val='t'))
    # Let each application lay out a native horizontal legend. Fixed manual
    # bounds (especially h=.06) cause WPS to suppress all three entries.
    node(legend, 'layout')
    node(legend, 'overlay', val=0)
    text_style(legend, 750, True)
    order = 'legendPos legendEntry layout overlay spPr txPr extLst'.split()
    legend[:] = sorted(legend, key=lambda child: order.index(E.QName(child).localname))
    layout = plot.find('c:layout/c:manualLayout', NS)
    if layout is not None:
        # Reserve space for title + legend above the highest data labels.
        for key, value in (('y', .30), ('h', .57)):
            field = layout.find('c:'+key, NS)
            if field is None:
                field = node(layout, key)
            field.set('val', str(value))
    return E.tostring(root)


def node(parent, name, **attrs):
    return E.SubElement(parent, '{'+C+'}'+name, {k:str(v) for k,v in attrs.items()})


def text_style(parent, size=900, bold=False, color='000000', family='Arial'):
    tx = node(parent, 'txPr')
    E.SubElement(tx, '{'+A+'}bodyPr', rot='0', wrap='square', vertOverflow='overflow')
    E.SubElement(tx, '{'+A+'}lstStyle')
    p = E.SubElement(tx, '{'+A+'}p')
    pr = E.SubElement(p, '{'+A+'}pPr')
    run = E.SubElement(pr, '{'+A+'}defRPr', sz=str(size), b='1' if bold else '0')
    E.SubElement(E.SubElement(run, '{'+A+'}solidFill'), '{'+A+'}srgbClr', val=color)
    E.SubElement(run, '{'+A+'}latin', typeface=family)


def label_text(parent, value, size, highlight=None):
    """Explicit native label text and per-label fill survive WPS rendering."""
    rich = node(node(parent, 'tx'), 'rich')
    E.SubElement(rich, '{'+A+'}bodyPr', wrap='none', lIns='0', rIns='0', tIns='0', bIns='0')
    E.SubElement(rich, '{'+A+'}lstStyle')
    paragraph = E.SubElement(rich, '{'+A+'}p')
    run = E.SubElement(paragraph, '{'+A+'}r')
    props = E.SubElement(run, '{'+A+'}rPr', sz=str(size), b='1')
    E.SubElement(E.SubElement(props, '{'+A+'}solidFill'), '{'+A+'}srgbClr', val='000000')
    if highlight:
        E.SubElement(E.SubElement(props, '{'+A+'}highlight'), '{'+A+'}srgbClr', val=highlight)
    E.SubElement(props, '{'+A+'}latin', typeface='Arial')
    E.SubElement(run, '{'+A+'}t').text = f'{value:.1%}'


def paint(parent, color=None, line=None, width=12700):
    sp = node(parent, 'spPr')
    E.SubElement(sp, '{'+A+'}noFill') if color is None else E.SubElement(E.SubElement(sp, '{'+A+'}solidFill'), '{'+A+'}srgbClr', val=color)
    ln = E.SubElement(sp, '{'+A+'}ln', w=str(width))
    E.SubElement(ln, '{'+A+'}noFill') if line is None else E.SubElement(E.SubElement(ln, '{'+A+'}solidFill'), '{'+A+'}srgbClr', val=line)
    return sp


class Values:
    def __init__(self, path):
        self.book = load_workbook(path, data_only=False)
        self.memo = {}

    def cell(self, sheet, address):
        address = address.replace('$','')
        key = sheet, address
        if key in self.memo:
            return self.memo[key]
        value = self.book[sheet][address].value
        # The old export created region-series references but left their
        # source cells empty. Supply weighted, formula-backed FTT totals.
        match=re.fullmatch(r'([A-Z]+)(36|37)',address)
        if sheet=='FTT month-quarter' and match and value is None:
            col,row=match.groups();first=0 if row=='36' else 2
            value=f'=({col}{9+first}+{col}{10+first}-{col}{17+first}-{col}{18+first})/({col}{2+first}+{col}{3+first})'
            self.book[sheet][address]=value
        if isinstance(value,str) and value.startswith('='):
            expression = value[1:]
            guarded = expression.startswith('IFERROR(')
            if guarded:
                expression = expression[8:-4]
            def sum_range(match):
                owner = match[1] or sheet
                return str(sum(v for v in self.range(owner,match[2]) if isinstance(v,(int,float))))
            expression = re.sub(r"SUM\((?:'([^']+)'!)?([A-Z$]+\d+:[A-Z$]+\d+)\)",sum_range,expression)
            expression = re.sub(r'\$?[A-Z]+\$?\d+',lambda m:str(self.cell(sheet,m[0]) or 0),expression)
            tree = ast.parse(expression,mode='eval')
            def calculate(n):
                if isinstance(n,ast.Expression): return calculate(n.body)
                if isinstance(n,ast.Constant) and isinstance(n.value,(int,float)): return n.value
                if isinstance(n,ast.BinOp):
                    x,y=calculate(n.left),calculate(n.right)
                    if isinstance(n.op,ast.Add): return x+y
                    if isinstance(n.op,ast.Sub): return x-y
                    if isinstance(n.op,ast.Div): return x/y
                    if isinstance(n.op,ast.Mult): return x*y
                raise ValueError('Unsupported report formula')
            try: value=calculate(tree)
            except ZeroDivisionError: value='' if guarded else None
        self.memo[key]=value
        return value

    def range(self,sheet,address):
        x1,y1,x2,y2=range_boundaries(address)
        return [self.cell(sheet,f'{get_column_letter(x)}{y}') for y in range(y1,y2+1) for x in range(x1,x2+1)]

    def reference(self,formula):
        if re.fullmatch(r'"(?:[^"]|"")*"', formula or ''):
            return [formula[1:-1].replace('""', '"')]
        sheet,address=formula.rsplit('!',1)
        return self.range(sheet.strip("'").replace("''","'"),address)


def styled_chart(data, metric, region, kind, values, defect_colors=None):
    from .template_output import TREND_FORMAT, RANK_FORMAT
    from report_tools.hfpa.agents.color_manager import ColorManager
    root=E.fromstring(data)
    chart=root.find('c:chart',NS)
    plot=chart.find('c:plotArea',NS)
    trend=kind=='trend'
    # Rebuild caches from the workbook formulas, never from screenshot numbers.
    for ref in root.findall('.//c:numRef',NS)+root.findall('.//c:strRef',NS):
        f=ref.find('c:f',NS)
        if f is None: continue
        numeric=E.QName(ref).localname=='numRef'
        points=values.reference(f.text)
        if E.QName(ref.getparent()).localname=='cat' and any(isinstance(v,str) for v in points):
            ref.tag='{'+C+'}strRef';numeric=False
        for old in list(ref):
            if old is not f: ref.remove(old)
        cache=node(ref,'numCache' if numeric else 'strCache')
        if numeric: node(cache,'formatCode').text=PERCENT_FORMAT
        node(cache,'ptCount',val=len(points))
        for i,value in enumerate(points):
            if value is None or value=='': continue
            if not trend and E.QName(ref.getparent()).localname == 'cat':
                value = '\n'.join(textwrap.wrap(str(value), width=24,
                                              break_long_words=True, break_on_hyphens=False))
            pt=node(cache,'pt',idx=i)
            node(pt,'v').text=str(value)
    if not trend:
        # WPS recalculates strRef from the unwrapped source cells and discards
        # cached line breaks. Literal categories in the Excel chart keep the
        # complete names, while value series still refer to their source cells.
        for category in plot.findall('.//c:ser/c:cat', NS):
            points = category.findall('.//c:pt', NS)
            labels = [(point.get('idx'), point.find('c:v', NS).text or '') for point in points]
            count = category.find('.//c:ptCount', NS)
            count = int(count.get('val')) if count is not None else max((int(index) for index, _ in labels), default=-1)+1
            category.clear()
            literal = node(category, 'strLit')
            node(literal, 'ptCount', val=count)
            for index, value in labels:
                node(node(literal, 'pt', idx=index), 'v').text = value
    for old in list(chart):
        if E.QName(old).localname in ('title','legend'): chart.remove(old)
    title=E.Element('{'+C+'}title')
    rich=node(node(title,'tx'),'rich')
    E.SubElement(rich,'{'+A+'}bodyPr', anchor='ctr', anchorCtr='1', wrap='none',
                 lIns='0', rIns='0', tIns='0', bIns='0')
    E.SubElement(rich,'{'+A+'}lstStyle')
    paragraph=E.SubElement(rich,'{'+A+'}p')
    E.SubElement(paragraph, '{'+A+'}pPr', algn='ctr')
    run=E.SubElement(paragraph,'{'+A+'}r')
    props=E.SubElement(run,'{'+A+'}rPr',sz='1050' if trend else '950',b='1')
    E.SubElement(E.SubElement(props, '{'+A+'}solidFill'), '{'+A+'}srgbClr', val='000000')
    E.SubElement(props,'{'+A+'}latin',typeface='Arial Black' if trend else 'Arial')
    heading=(f'{"VN" if region=="VN" else "Indo"} {metric}% TREND TRACKING' if trend else f'TOP {metric} MODELS' if kind=='model' else 'TOP DEFECTS OF @TOP1 MODEL')
    E.SubElement(run,'{'+A+'}t').text=heading
    node(title,'overlay',val=0); chart.insert(0,title)
    if not trend:
        # Use the native centred title position. Manual factor positions make
        # WPS shift short titles to the right; the plot reserves the top band.
        node(title, 'layout')
    for old in plot.findall('c:layout',NS):plot.remove(old)
    layout=E.Element('{'+C+'}layout'); manual=node(layout,'manualLayout')
    node(manual,'layoutTarget',val='inner')
    for k,v in dict(xMode='factor' if trend else 'edge',yMode='factor' if trend else 'edge',
                   wMode='factor',hMode='factor').items():node(manual,k,val=v)
    # The TOP title owns the upper band. Leave a gap for outside-end numbers
    # before the bars, and retain the lower band for complete model names.
    for k,v in dict(x=.07 if trend else RANK_PLOT_X,y=.30 if trend else .36,
                   w=.90 if trend else RANK_PLOT_WIDTH,h=.57 if trend else .28).items():node(manual,k,val=v)
    plot.insert(0,layout)
    for old in root.findall('c:spPr',NS):root.remove(old)
    paint(root,'FFFFFF','DDDDDD' if trend else None,12700)
    series=plot.findall('.//c:ser',NS)
    trend_points = [{int(p.get('idx')): float(p.find('c:v', NS).text)
                     for p in ser.findall('c:val/c:numRef/c:numCache/c:pt', NS)}
                    for ser in series] if trend else []
    label_positions = {}
    for index in sorted({index for points in trend_points for index in points}):
        # At each month the lowest label goes below, the highest above and
        # the middle in the centre, even when factory lines cross or coincide.
        ranked = sorted(((points[index], i) for i, points in enumerate(trend_points) if index in points),
                        key=lambda item: (item[0], 1 if item[1] == 0 else 0 if item[1] == 1 else 2))
        for rank, (_, owner) in enumerate(ranked):
            label_positions[owner, index] = ('b' if rank == 0 else
                                             't' if rank == len(ranked)-1 else 'ctr')
    colors = [colour.lstrip('#') for colour in TREND_FORMAT[metric, region][-2]]
    for i,ser in enumerate(series):
        for old in list(ser):
            if E.QName(old).localname in ('spPr','marker','dLbls','dPt','smooth'):ser.remove(old)
        if trend:
            sp=paint(ser,None,colors[i],19050 if i==0 else 12700);ser.remove(sp);ser.insert(3,sp)
            # Point markers can overprint nearby percentage labels in WPS.
            # Native coloured lines and legend keys identify each factory.
            marker=E.Element('{'+C+'}marker');node(marker,'symbol',val='none');ser.insert(4,marker)
        else:
            labels = {int(point.get('idx')): point.find('c:v', NS).text or ''
                      for point in ser.findall('c:cat//c:pt', NS)}
            for j, label in labels.items():
                color = (RANK_FORMAT[metric, region, False][-1] if kind == 'model' else
                         (defect_colors or {})[ColorManager.normalize_name(label)])
                color = color.lstrip('#')
                point=E.Element('{'+C+'}dPt');node(point,'idx',val=j);paint(point,color);ser.insert(3+j,point)
        labels=E.Element('{'+C+'}dLbls')
        numeric = (trend_points[i] if trend else
                   {int(p.get('idx')): float(p.find('c:v', NS).text)
                    for p in ser.findall('c:val/c:numRef/c:numCache/c:pt', NS)})
        for index, value in numeric.items():
            point=node(labels,'dLbl');node(point,'idx',val=index)
            highlight = TREND_FORMAT[metric, region][-1].lstrip('#') if trend and i == 0 else None
            label_text(point, value, 800 if trend else 700, highlight)
            node(point,'numFmt',formatCode=PERCENT_FORMAT,sourceLinked=0)
            if trend:
                paint(point, highlight or 'FFFFFF')
                text_style(point, 800, True)
            node(point,'dLblPos',val=label_positions[i,index] if trend else 'outEnd')
            node(point,'showLegendKey',val=0);node(point,'showVal',val=0);node(point,'showCatName',val=0);node(point,'showSerName',val=0);node(point,'showPercent',val=0)
        node(labels,'numFmt',formatCode=PERCENT_FORMAT,sourceLinked=0)
        text_style(labels,750 if trend else 700,True)
        node(labels,'dLblPos',val='ctr' if trend and i==0 else 'b' if trend and i==2 else 't' if trend else 'outEnd')
        node(labels,'showLegendKey',val=0);node(labels,'showVal',val=1)
        node(labels,'showCatName',val=0);node(labels,'showSerName',val=0)
        pos=next((j for j,n in enumerate(ser) if E.QName(n).localname in ('cat','val')),len(ser));ser.insert(pos,labels)
        if trend: node(ser,'smooth',val=0)
    for parent in list(plot):
        if E.QName(parent).localname in ('lineChart','barChart'):
            for old in parent.findall('c:dLbls',NS):parent.remove(old)
            if not trend:
                gap=parent.find('c:gapWidth',NS)
                if gap is not None:gap.set('val','160')
    for axis in plot.findall('c:catAx',NS)+plot.findall('c:valAx',NS):
        vertical=E.QName(axis).localname=='valAx'
        for old in list(axis):
            if E.QName(old).localname in ('title','txPr','spPr','majorGridlines'):axis.remove(old)
        delete=axis.find('c:delete',NS)
        if delete is None:delete=node(axis,'delete',val=0)
        delete.set('val','1' if vertical and not trend else '0')
        text_style(axis,650 if vertical or not trend else 750,not vertical,'8F8F8F' if vertical else '000000')
        if not vertical:
            tick = axis.find('c:tickLblPos', NS)
            if tick is not None: tick.set('val', 'low')
            offset = axis.find('c:lblOffset', NS)
            if offset is not None: offset.set('val', '100')
            if not trend:
                for tag in ('tickLblSkip', 'tickMarkSkip'):
                    element = axis.find('c:'+tag, NS)
                    if element is None: element=node(axis,tag)
                    element.set('val', '1')
        paint(axis,None,'EEEEEE' if not vertical else None,6350)
        if vertical and trend:
            number_format=axis.find('c:numFmt',NS)
            if number_format is None: number_format=node(axis,'numFmt')
            number_format.set('formatCode',PERCENT_FORMAT);number_format.set('sourceLinked','0')
            scaling=axis.find('c:scaling',NS)
            for old in list(scaling):
                if E.QName(old).localname in ('min','max'):scaling.remove(old)
            low,high,step=(.5,.85,.05) if metric=='FTT' else (.80,.90,.02) if region=='VN' else (.70,.98,.04)
            observations=[v for points in trend_points for v in points.values() if math.isfinite(v)]
            if observations:
                low=max(0, min(low, math.floor((min(observations)-step)/step)*step))
                high=max(high, math.ceil((max(observations)+step)/step)*step)
            if high-low > step*10: step=math.ceil((high-low)/.10)*.01
            node(scaling,'max',val=high);node(scaling,'min',val=low)
            unit=axis.find('c:majorUnit',NS)
            if unit is None:unit=node(axis,'majorUnit')
            unit.set('val',str(step)); grid=node(axis,'majorGridlines');paint(grid,None,'F1F1F1',6350)
    if trend:
        legend=node(chart,'legend');node(legend,'legendPos',val='t');node(legend,'layout')
        node(legend,'overlay',val=0);text_style(legend,800,True)
    ordered={
        'chart':'title autoTitleDeleted pivotFmts view3D floor sideWall backWall plotArea legend plotVisOnly dispBlanksAs showDLblsOverMax extLst',
        'catAx':'axId scaling delete axPos majorGridlines minorGridlines title numFmt majorTickMark minorTickMark tickLblPos spPr txPr crossAx crosses crossesAt auto lblAlgn lblOffset tickLblSkip tickMarkSkip noMultiLvlLbl extLst',
        'valAx':'axId scaling delete axPos majorGridlines minorGridlines title numFmt majorTickMark minorTickMark tickLblPos spPr txPr crossAx crosses crossesAt crossBetween majorUnit minorUnit dispUnits extLst',
        'ser':'idx order tx spPr invertIfNegative pictureOptions marker dPt dLbls trendline errBars cat val smooth extLst',
        'legend':'legendPos legendEntry layout overlay spPr txPr extLst',
        'title':'tx layout overlay spPr txPr extLst',
        'chartSpace':'date1904 lang roundedCorners AlternateContent style clrMapOvr pivotSource protection chart spPr txPr externalData printSettings userShapes extLst',
    }
    for element in root.iter():
        order=ordered.get(E.QName(element).localname)
        if order:
            keys=order.split();element[:]=sorted(element,key=lambda n:keys.index(E.QName(n).localname) if E.QName(n).localname in keys else 999)
    return compatible_trend_legend(E.tostring(root)) if trend else E.tostring(root)


def prepare_excel_charts(xlsx):
    """Finish and save the Excel charts before PowerPoint reads any chart."""
    from .template_output import copy_package, repair_excel_chart_references, defect_colors_for
    values = Values(xlsx)
    workbook_parts = {}
    try:
        with ZipFile(xlsx) as book:
            sources = excel_chart_sources(book)
            defect_names = []
            for (_, _, kind), name in sources.items():
                if kind != 'defect':
                    continue
                root = E.fromstring(book.read(name))
                for category in root.findall('.//c:ser/c:cat', NS):
                    formula = category.find('.//c:f', NS)
                    if formula is not None:
                        defect_names.extend(values.reference(formula.text))
                    else:
                        defect_names.extend(category.xpath('.//c:pt/c:v/text()', namespaces=NS))
            defect_colors = defect_colors_for(defect_names)
            for (metric, region, kind), name in sources.items():
                workbook_parts[name] = styled_chart(book.read(name), metric, region, kind, values, defect_colors)
        # Keep the source formulas and populate their cached results for consumers
        # which display embedded workbooks before recalculating them.
        with ZipFile(xlsx) as book:
            for index,sheet in enumerate(values.book,1):
                name=f'xl/worksheets/sheet{index}.xml'
                xml=E.fromstring(book.read(name));x='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
                for cell in xml.iter(x+'c'):
                    if cell.find(x+'f') is None:continue
                    value=values.cell(sheet.title,cell.get('r'))
                    cached=cell.find(x+'v')
                    if cached is None:cached=E.SubElement(cell,x+'v')
                    cached.text='' if value is None else str(value)
                    if isinstance(value,str):cell.set('t','str')
                    else:cell.attrib.pop('t',None)
                if sheet.title=='FTT month-quarter':
                    data=xml.find(x+'sheetData')
                    for row in (36,37):
                        record=next((r for r in data if r.get('r')==str(row)),None)
                        if record is None:record=E.SubElement(data,x+'row',r=str(row));data[:]=sorted(data,key=lambda e:int(e.get('r')))
                        columns={re.fullmatch(r'([A-Z]+)'+str(row),address)[1]
                                 for owner,address in values.memo
                                 if owner==sheet.title and re.fullmatch(r'([A-Z]+)'+str(row),address)}
                        for col in sorted(columns,key=lambda name:range_boundaries(name+'1')[0]):
                            address=f'{col}{row}';formula=sheet[address].value
                            if isinstance(formula,str) and formula.startswith('='):
                                old=next((c for c in record if c.get('r')==address),None)
                                cell=old if old is not None else E.SubElement(record,x+'c',r=address)
                                for child in list(cell):cell.remove(child)
                                E.SubElement(cell,x+'f').text=formula[1:]
                                result=values.cell(sheet.title,address)
                                E.SubElement(cell,x+'v').text='' if result is None else str(result)
                workbook_parts[name]=E.tostring(xml)
        buffer = BytesIO()
        copy_package(xlsx, buffer, workbook_parts)
        Path(xlsx).write_bytes(buffer.getvalue())
        repair_excel_chart_references(xlsx)
    finally:
        values.book.close()


def chart_rectangle(slide_width, slide_height, region, kind):
    """Disjoint chart slots with a gutter for the existing down arrow."""
    top = .145 if region == 'VN' else .590
    if kind == 'trend':
        box = (.035, top, .615, .375)
    elif kind == 'model':
        box = (.670, top, .315, .157)
    else:
        box = (.670, top + .207, .315, .168)
    x, y, w, h = box
    return tuple(round(value * dimension) for value, dimension in
                 zip((x, y, w, h), (slide_width, slide_height, slide_width, slide_height)))


def convert(pptx, xlsx, output_pptx, output_xlsx=None):
    """Use native Office Copy/Paste; Excel owns the complete chart design."""
    from .office_charts import copy_charts
    return copy_charts(pptx, xlsx, output_pptx, output_xlsx)
