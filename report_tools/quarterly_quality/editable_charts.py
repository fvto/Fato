"""Transfer Excel charts into QBR picture slots without rebuilding the deck."""
import ast
from copy import deepcopy
from io import BytesIO
from pathlib import Path
import re
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


def node(parent, name, **attrs):
    return E.SubElement(parent, '{'+C+'}'+name, {k:str(v) for k,v in attrs.items()})


def text_style(parent, size=900, bold=False, color='000000', family='Arial'):
    tx = node(parent, 'txPr')
    E.SubElement(tx, '{'+A+'}bodyPr')
    E.SubElement(tx, '{'+A+'}lstStyle')
    p = E.SubElement(tx, '{'+A+'}p')
    pr = E.SubElement(p, '{'+A+'}pPr')
    run = E.SubElement(pr, '{'+A+'}defRPr', sz=str(size), b='1' if bold else '0')
    E.SubElement(E.SubElement(run, '{'+A+'}solidFill'), '{'+A+'}srgbClr', val=color)
    E.SubElement(run, '{'+A+'}latin', typeface=family)


def paint(parent, color=None, line=None, width=12700):
    sp = node(parent, 'spPr')
    E.SubElement(sp, '{'+A+'}noFill') if color is None else E.SubElement(E.SubElement(sp, '{'+A+'}solidFill'), '{'+A+'}srgbClr', val=color)
    ln = E.SubElement(sp, '{'+A+'}ln', w=str(width))
    E.SubElement(ln, '{'+A+'}noFill') if line is None else E.SubElement(E.SubElement(ln, '{'+A+'}solidFill'), '{'+A+'}srgbClr', val=line)
    return sp


def slide_text(tree, ident, text, x, y, w, h, size=1000, bold=True, color='000000', family='Arial'):
    shape=E.SubElement(tree,'{'+P+'}sp')
    nv=E.SubElement(shape,'{'+P+'}nvSpPr');E.SubElement(nv,'{'+P+'}cNvPr',id=str(ident),name=text);E.SubElement(nv,'{'+P+'}cNvSpPr',txBox='1');E.SubElement(nv,'{'+P+'}nvPr')
    sp=E.SubElement(shape,'{'+P+'}spPr');xf=E.SubElement(sp,'{'+A+'}xfrm');E.SubElement(xf,'{'+A+'}off',x=str(int(x)),y=str(int(y)));E.SubElement(xf,'{'+A+'}ext',cx=str(int(w)),cy=str(int(h)))
    E.SubElement(sp,'{'+A+'}noFill');E.SubElement(E.SubElement(sp,'{'+A+'}ln'),'{'+A+'}noFill')
    tx=E.SubElement(shape,'{'+P+'}txBody');E.SubElement(tx,'{'+A+'}bodyPr',wrap='square',lIns='0',rIns='0',tIns='0',bIns='0',anchor='ctr');E.SubElement(tx,'{'+A+'}lstStyle')
    paragraph=E.SubElement(tx,'{'+A+'}p');E.SubElement(paragraph,'{'+A+'}pPr',algn='ctr')
    for j,line in enumerate(text.split('\n')):
        if j:E.SubElement(paragraph,'{'+A+'}br')
        run=E.SubElement(paragraph,'{'+A+'}r');pr=E.SubElement(run,'{'+A+'}rPr',sz=str(size),b='1' if bold else '0')
        E.SubElement(E.SubElement(pr,'{'+A+'}solidFill'),'{'+A+'}srgbClr',val=color);E.SubElement(pr,'{'+A+'}latin',typeface=family);E.SubElement(run,'{'+A+'}t').text=line


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
        sheet,address=formula.rsplit('!',1)
        return self.range(sheet.strip("'").replace("''","'"),address)


def styled_chart(data, metric, region, kind, values):
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
        if numeric: node(cache,'formatCode').text='0.0%'
        node(cache,'ptCount',val=len(points))
        for i,value in enumerate(points):
            if value is None or value=='': continue
            pt=node(cache,'pt',idx=i)
            node(pt,'v').text=str(value)
    for old in list(chart):
        if E.QName(old).localname in ('title','legend'): chart.remove(old)
    title=E.Element('{'+C+'}title')
    rich=node(node(title,'tx'),'rich')
    E.SubElement(rich,'{'+A+'}bodyPr'); E.SubElement(rich,'{'+A+'}lstStyle')
    paragraph=E.SubElement(rich,'{'+A+'}p'); run=E.SubElement(paragraph,'{'+A+'}r')
    props=E.SubElement(run,'{'+A+'}rPr',sz='1200' if trend else '1300',b='1')
    E.SubElement(props,'{'+A+'}latin',typeface='Arial Black' if trend else 'Arial')
    heading=(f'{"VN" if region=="VN" else "Indo"} {metric}% TREND TRACKING' if trend else f'TOP {metric} MODELS' if kind=='model' else 'TOP DEFECTS OF @TOP1 MODEL')
    E.SubElement(run,'{'+A+'}t').text=heading
    node(title,'overlay',val=0); chart.insert(0,title)
    if kind=='defect':
        tl=node(title,'layout');tm=node(tl,'manualLayout')
        for k,v in dict(xMode='factor',yMode='factor',wMode='factor',hMode='factor',x=.25,y=.02,w=.73,h=.15).items():node(tm,k,val=v)
    for old in plot.findall('c:layout',NS):plot.remove(old)
    layout=E.Element('{'+C+'}layout'); manual=node(layout,'manualLayout')
    node(manual,'layoutTarget',val='inner')
    for k,v in dict(xMode='factor',yMode='factor',wMode='factor',hMode='factor').items():node(manual,k,val=v)
    for k,v in dict(x=.05 if trend else .02,y=.17 if trend else .30,w=.93 if trend else .96,h=.70 if trend else .32).items():node(manual,k,val=v)
    plot.insert(0,layout)
    for old in root.findall('c:spPr',NS):root.remove(old)
    paint(root,'FFFFFF','DDDDDD' if trend else None,12700)
    series=plot.findall('.//c:ser',NS)
    aggregate={int(p.get('idx')):float(p.find('c:v',NS).text) for p in series[0].findall('c:val/c:numRef/c:numCache/c:pt',NS)} if trend else {}
    colors=({'HFPA':{'VN':['4B3768','DBC8FF','8D698C'],'Indo':['4B3768','D8C4D6','B7E4FF']},'FTT':{'VN':['082567','CAA367','8497B0'],'Indo':['6550BD','946BDB','72CAC4']}}[metric][region])
    for i,ser in enumerate(series):
        for old in list(ser):
            if E.QName(old).localname in ('spPr','marker','dLbls','dPt','smooth'):ser.remove(old)
        if trend:
            sp=paint(ser,None,colors[i],19050 if i==0 else 12700);ser.remove(sp);ser.insert(3,sp)
            marker=E.Element('{'+C+'}marker');node(marker,'symbol',val='circle');node(marker,'size',val=4);paint(marker,'FFFFFF',colors[i],9525);ser.insert(4,marker)
        else:
            palette=(['5B9BD5']*3 if region=='VN' else ['9280AA']*3) if kind=='model' else (['5082BE','10263F','50672D'] if metric=='HFPA' and region=='VN' else ['50672D','10263F','5082BE'] if metric=='HFPA' else ['50672D','10263F','FA9440'])
            for j,color in enumerate(palette):
                point=E.Element('{'+C+'}dPt');node(point,'idx',val=j);paint(point,color);ser.insert(3+j,point)
        labels=E.Element('{'+C+'}dLbls')
        if trend:
            numeric={int(p.get('idx')):float(p.find('c:v',NS).text) for p in ser.findall('c:val/c:numRef/c:numCache/c:pt',NS)}
            for index in range(3):
                point=node(labels,'dLbl');node(point,'idx',val=index)
                above=numeric.get(index,0)>=aggregate.get(index,0)
                if i:
                    layout=node(point,'layout');manual=node(layout,'manualLayout');node(manual,'x',val=0);node(manual,'y',val=-.045 if above else .045)
                node(point,'dLblPos',val='ctr' if i==0 else 't' if above else 'b')
                node(point,'showLegendKey',val=0);node(point,'showVal',val=1);node(point,'showCatName',val=0);node(point,'showSerName',val=0);node(point,'showPercent',val=0)
        node(labels,'numFmt',formatCode='0.0%',sourceLinked=0)
        if trend and i==0:paint(labels,'00FFFF' if region=='VN' else 'FFFF4D')
        text_style(labels,800,True)
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
        text_style(axis,700 if vertical else 900,not vertical,'BFBFBF' if vertical else '000000')
        paint(axis,None,'EEEEEE' if not vertical else None,6350)
        if vertical and trend:
            number_format=axis.find('c:numFmt',NS)
            if number_format is not None:number_format.set('formatCode','0%')
            scaling=axis.find('c:scaling',NS)
            for old in list(scaling):
                if E.QName(old).localname in ('min','max'):scaling.remove(old)
            low,high,step=(.5,.85,.05) if metric=='FTT' else (.80,.90,.02) if region=='VN' else (.70,.98,.04)
            node(scaling,'max',val=high);node(scaling,'min',val=low)
            unit=axis.find('c:majorUnit',NS)
            if unit is None:unit=node(axis,'majorUnit')
            unit.set('val',str(step)); grid=node(axis,'majorGridlines');paint(grid,None,'F1F1F1',6350)
    if trend:
        legend=node(chart,'legend');node(legend,'legendPos',val='t');ll=node(legend,'layout');lm=node(ll,'manualLayout')
        for k,v in dict(xMode='factor',yMode='factor',wMode='factor',hMode='factor',x=.74,y=.02,w=.25,h=.06).items():node(lm,k,val=v)
        node(legend,'overlay',val=1);text_style(legend,700,True)
    ordered={
        'chart':'title autoTitleDeleted pivotFmts view3D floor sideWall backWall plotArea legend plotVisOnly dispBlanksAs showDLblsOverMax extLst',
        'catAx':'axId scaling delete axPos majorGridlines minorGridlines title numFmt majorTickMark minorTickMark tickLblPos spPr txPr crossAx crosses crossesAt auto lblAlgn lblOffset tickLblSkip tickMarkSkip noMultiLvlLbl extLst',
        'valAx':'axId scaling delete axPos majorGridlines minorGridlines title numFmt majorTickMark minorTickMark tickLblPos spPr txPr crossAx crosses crossesAt crossBetween majorUnit minorUnit dispUnits extLst',
        'ser':'idx order tx spPr invertIfNegative pictureOptions marker dPt dLbls trendline errBars cat val smooth extLst',
        'legend':'legendPos legendEntry layout overlay spPr txPr extLst',
    }
    for element in root.iter():
        order=ordered.get(E.QName(element).localname)
        if order:
            keys=order.split();element[:]=sorted(element,key=lambda n:keys.index(E.QName(n).localname) if E.QName(n).localname in keys else 999)
    return E.tostring(root)


def convert(pptx, xlsx, output_pptx, output_xlsx=None):
    """Replace only the 12 chart pictures, preserving every other slide part."""
    from .template_output import copy_package, repair_excel_chart_references
    values=Values(xlsx)
    workbook_parts={}
    parts={}
    with ZipFile(pptx) as deck, ZipFile(xlsx) as book:
        content=E.fromstring(deck.read('[Content_Types].xml'))
        existing={e.get('PartName') for e in content}
        count=0
        for slide_number,metric,mapping in [(7,'FTT',[(2,8,'VN','trend'),(5,9,'Indo','trend'),(6,10,'VN','model'),(3,11,'Indo','model'),(7,12,'VN','defect'),(13,13,'Indo','defect')]),(8,'HFPA',[(2,2,'VN','trend'),(3,3,'Indo','trend'),(8,4,'VN','model'),(11,5,'Indo','model'),(10,6,'VN','defect'),(12,7,'Indo','defect')])]:
            name=f'ppt/slides/slide{slide_number}.xml'
            # Full and compact QBR templates share chart-slot IDs, not slide numbers.
            candidates=[n for n in deck.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n) and f'QMS {metric} ANALYSIS' in ''.join(E.fromstring(deck.read(n)).xpath('.//a:t/text()',namespaces=NS))]
            if len(candidates)!=1:raise ValueError(f'Expected one QMS {metric} slide.')
            name=candidates[0];slide_number=int(re.search(r'slide(\d+)',name)[1])
            slide=E.fromstring(deck.read(name))
            if f'QMS {metric} ANALYSIS' not in ''.join(slide.xpath('.//a:t/text()',namespaces=NS)):continue
            rel_name=f'ppt/slides/_rels/slide{slide_number}.xml.rels';rels=E.fromstring(deck.read(rel_name))
            for shape_id,chart_index,region,kind in mapping:
                pictures=slide.xpath(f'.//p:pic[p:nvPicPr/p:cNvPr[@id="{shape_id}"]]',namespaces=NS)
                if not pictures:continue
                pic=pictures[0];count+=1
                source_name=f'xl/charts/chart{chart_index}.xml'
                styled=styled_chart(book.read(source_name),metric,region,kind,values)
                workbook_parts[source_name]=styled
                root=E.fromstring(styled)
                # Keep headings and legends in reserved editable slide text
                # areas. Native chart auto-layout otherwise competes with the
                # narrow arrow/title gutters in this exact QBR template.
                ch=root.find('c:chart',NS)
                for old in ch.findall('c:title',NS)+ch.findall('c:legend',NS):ch.remove(old)
                auto=ch.find('c:autoTitleDeleted',NS)
                if auto is None:auto=E.Element('{'+C+'}autoTitleDeleted');ch.insert(0,auto)
                auto.set('val','1')
                xf=pic.find('p:spPr/a:xfrm',NS)
                off=xf.find('a:off',NS);ext=xf.find('a:ext',NS)
                x,y,w,h=(int(off.get('x')),int(off.get('y')),int(ext.get('cx')),int(ext.get('cy')))
                tree=pic.getparent()
                heading=(f'{"VN" if region=="VN" else "Indo"} {metric}% TREND TRACKING' if kind=='trend' else f'TOP {metric} MODELS' if kind=='model' else 'TOP DEFECTS OF @TOP1 MODEL')
                slide_text(tree,1000+count*10,heading,x+w*(.12 if kind=='trend' else .23 if kind=='defect' else 0),y+60000,w*(.76 if kind=='trend' else .75 if kind=='defect' else 1),230000,1100 if kind=='trend' else 1400 if kind=='model' else 1000,family='Arial Black' if kind=='trend' else 'Arial')
                if kind=='trend':
                    legend_names=['VietNam','VH','VH2'] if region=='VN' else ['Indo','JV','JV2']
                    legend_colors=({'HFPA':{'VN':['4B3768','DBC8FF','8D698C'],'Indo':['4B3768','D8C4D6','B7E4FF']},'FTT':{'VN':['082567','CAA367','8497B0'],'Indo':['6550BD','946BDB','72CAC4']}}[metric][region])
                    for j,label in enumerate(legend_names):slide_text(tree,1001+count*10+j,'—○ '+label,x+w*.72+j*w*.09,y+60000,w*.09,200000,650,True,legend_colors[j])
                else:
                    for axis in root.findall('.//c:catAx',NS):
                        axis.find('c:delete',NS).set('val','1')
                        tick=axis.find('c:tickLblPos',NS)
                        if tick is not None:tick.set('val','none')
                    cats=root.findall('.//c:ser/c:cat//c:pt/c:v',NS)
                    for j,cat in enumerate(cats):
                        label=cat.text or ''
                        # Wrap long names into their own equally sized gutters.
                        if len(label)>24:
                            words=label.split();lines=['']
                            for word in words:
                                if len(lines[-1])+len(word)>20:lines.append(word)
                                else:lines[-1]=(lines[-1]+' '+word).strip()
                            label='\n'.join(lines)
                        slide_text(tree,1001+count*10+j,label,x+j*w/3+40000,y+h*.79,w/3-80000,h*.20,800,True)
                external=node(root,'externalData',**{'{'+R+'}id':'rIdWorkbook'});node(external,'autoUpdate',val=0)
                chart_name=f'ppt/charts/qbr{count}.xml';parts[chart_name]=E.tostring(root)
                chart_rels=E.Element('{'+PK+'}Relationships',nsmap={None:PK})
                E.SubElement(chart_rels,'{'+PK+'}Relationship',Id='rIdWorkbook',Type=R+'/package',Target='../embeddings/qbr-source.xlsx')
                parts[f'ppt/charts/_rels/qbr{count}.xml.rels']=E.tostring(chart_rels)
                rid=f'rIdEditableQBR{count}'
                E.SubElement(rels,'{'+PK+'}Relationship',Id=rid,Type=R+'/chart',Target=f'../charts/qbr{count}.xml')
                frame=E.Element('{'+P+'}graphicFrame')
                nv=E.SubElement(frame,'{'+P+'}nvGraphicFramePr');E.SubElement(nv,'{'+P+'}cNvPr',id=str(shape_id),name=f'{metric} {region} {kind}');E.SubElement(nv,'{'+P+'}cNvGraphicFramePr');E.SubElement(nv,'{'+P+'}nvPr')
                transform=deepcopy(pic.find('p:spPr/a:xfrm',NS));transform.tag='{'+P+'}xfrm';frame.append(transform)
                graphic=E.SubElement(frame,'{'+A+'}graphic');data=E.SubElement(graphic,'{'+A+'}graphicData',uri=C);E.SubElement(data,'{'+C+'}chart',{'{'+R+'}id':rid})
                pic.getparent().replace(pic,frame)
                for relationship in list(rels):
                    if relationship.get('Id')==f'rIdHlink{shape_id}':rels.remove(relationship)
                E.SubElement(content,'{'+CT+'}Override',PartName='/'+chart_name,ContentType='application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
            parts[name]=E.tostring(slide);parts[rel_name]=E.tostring(rels)
            tree=slide.find('p:cSld/p:spTree',NS)
            for sid in ([8,12] if metric=='FTT' else [21,22]):
                arrow=slide.xpath(f'.//*[p:cNvPr[@id="{sid}"]]/..',namespaces=NS)
                if arrow and arrow[0].getparent() is tree:
                    tree.remove(arrow[0]);tree.append(arrow[0])
            parts[name]=E.tostring(slide)
        if count!=12:raise ValueError(f'Expected 12 QBR chart pictures; found {count}.')
        if not any(e.get('Extension')=='xlsx' for e in content):E.SubElement(content,'{'+CT+'}Default',Extension='xlsx',ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        parts['[Content_Types].xml']=E.tostring(content)
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
                            if old is not None:record.remove(old)
                            cell=E.SubElement(record,x+'c',r=address);E.SubElement(cell,x+'f').text=formula[1:];E.SubElement(cell,x+'v').text=str(values.cell(sheet.title,address))
            workbook_parts[name]=E.tostring(xml)
    buffer=BytesIO();copy_package(xlsx,buffer,workbook_parts);embedded=buffer.getvalue()
    if output_xlsx:
        Path(output_xlsx).write_bytes(embedded)
        repair_excel_chart_references(output_xlsx)
        embedded=Path(output_xlsx).read_bytes()
    parts['ppt/embeddings/qbr-source.xlsx']=embedded
    copy_package(pptx,output_pptx,parts)
    values.book.close()
    return count
