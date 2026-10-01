from pathlib import Path
from html import escape

def text(v):
    if v is None: return 'Chưa đủ dữ liệu'
    if isinstance(v,(int,float)): return f'{v:,.4f}'.rstrip('0').rstrip('.')
    return str(v)
def export_report(path:Path,body:dict):
    cols=body['columns'];headers=[body['headers'][i] for i in cols];rows=[[r[i] for i in cols] for r in body['rows']]
    meta=f"Dự án: {body.get('project','')} | Nghiệp vụ: {body['business']} | Quy tắc v{body.get('ruleVersion','')} | Mẫu v{body.get('templateVersion','')}"
    basis='Chênh lệch = giá chào − giá soát; % = chênh lệch / giá soát × 100, để trống khi giá soát bằng 0. Thành tiền và ảnh hưởng chỉ tính khi cách tính được xác nhận.'
    if path.suffix=='.xlsx':
        from openpyxl import Workbook
        from openpyxl.styles import Font,PatternFill,Alignment
        from openpyxl.utils import get_column_letter
        w=Workbook();s=w.active;s.title='Báo cáo CCM'
        def safe(v): return "'"+v if isinstance(v,str) and v.startswith(('=','+','-','@')) else v
        s.append([body['title']]);s.append([meta]);s.append([basis]);
        for section in body.get('sections',[]): s.append([safe(section)])
        s.append(headers);hr=s.max_row
        for row in rows: s.append([safe(v) for v in row])
        for cell in s[hr]: cell.fill=PatternFill('solid',fgColor='12396B');cell.font=Font(color='FFFFFF',bold=True);cell.alignment=Alignment(wrap_text=True,vertical='center')
        for i,col in enumerate(cols,1):
            s.column_dimensions[get_column_letter(i)].width=44 if col in [1,13,14] else 20
            for cells in s.iter_rows(min_row=hr+1,min_col=i,max_col=i):
                cell=cells[0];cell.alignment=Alignment(wrap_text=True,vertical='top');cell.font=Font(name='Calibri',size=11)
                if isinstance(cell.value,(float,int)): cell.number_format='0.00' if col==9 else '#,##0.00'
        for r in range(hr+1,s.max_row+1): s.row_dimensions[r].height=60
        s.row_dimensions[hr].height=44;s['A1'].font=Font(name='Calibri',size=20,bold=True,color='12396B')
        s.freeze_panes=f'A{hr+1}';s.auto_filter.ref=f'A{hr}:{get_column_letter(len(cols))}{s.max_row}'
        s.sheet_properties.pageSetUpPr.fitToPage=True;s.page_setup.orientation='landscape';s.page_setup.paperSize=s.PAPERSIZE_A3;s.page_setup.fitToWidth=1;s.page_setup.fitToHeight=0;s.print_title_rows=f'{hr}:{hr}'
        w.save(path);return
    if path.suffix=='.docx':
        from docx import Document
        from docx.shared import Inches,Pt
        from docx.enum.section import WD_ORIENT
        d=Document();section=d.sections[0];section.orientation=WD_ORIENT.LANDSCAPE;section.page_width=Inches(16.54);section.page_height=Inches(11.69);section.left_margin=section.right_margin=Inches(.5)
        d.styles['Normal'].font.name='Calibri';d.styles['Normal'].font.size=Pt(10)
        d.add_heading(body['title'],0);d.add_paragraph(meta);d.add_paragraph(basis)
        for item in body.get('sections',[]): d.add_heading(str(item),2)
        t=d.add_table(rows=1,cols=len(headers));t.style='Light Shading Accent 1'
        for c,h in zip(t.rows[0].cells,headers): c.text=h
        for row in rows:
            for c,v in zip(t.add_row().cells,row):c.text=text(v)
        from docx.oxml import OxmlElement
        repeat=OxmlElement('w:tblHeader');t.rows[0]._tr.get_or_add_trPr().append(repeat)
        d.save(path);return
    from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle
    from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
    from reportlab.lib.pagesizes import A3,landscape
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import os
    fonts=Path(os.environ.get('WINDIR','C:/Windows'))/'Fonts'
    font=fonts/'arial.ttf';bold=fonts/'arialbd.ttf'
    if not font.exists(): raise RuntimeError('Thiếu Arial Unicode để xuất PDF tiếng Việt')
    pdfmetrics.registerFont(TTFont('CCM',str(font)));pdfmetrics.registerFont(TTFont('CCM-Bold',str(bold)))
    styles=getSampleStyleSheet();p=ParagraphStyle('CCM',fontName='CCM',fontSize=8,leading=11);title=ParagraphStyle('CCM-title',fontName='CCM-Bold',fontSize=18,leading=22)
    flow=[Paragraph(escape(body['title']),title),Spacer(1,10),Paragraph(escape(meta),p),Paragraph(escape(basis),p),Spacer(1,12)]
    for section in body.get('sections',[]): flow.extend([Paragraph(escape(str(section)),p),Spacer(1,4)])
    width=landscape(A3)[0]-48
    weights=[3 if c in [1,13,14] else 1.2 for c in cols];sizes=[width*x/sum(weights) for x in weights]
    table=Table([[Paragraph(escape(text(v)),p) for v in row] for row in [headers]+rows],colWidths=sizes,repeatRows=1,hAlign='LEFT')
    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e3edf7')),('GRID',(0,0),(-1,-1),.3,colors.HexColor('#9db0c4')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5)]))
    flow.append(table);SimpleDocTemplate(str(path),pagesize=landscape(A3),leftMargin=24,rightMargin=24,topMargin=24,bottomMargin=24).build(flow)
