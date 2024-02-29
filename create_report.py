import csv
from reportlab.lib import colors
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import letter, inch, landscape
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_JUSTIFY, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image
from reportlab.platypus import Paragraph, PageBreak
from reportlab.pdfbase import pdfmetrics      
from reportlab.pdfbase.ttfonts import TTFont   
from reportlab.lib import pdfencrypt
from reportlab.lib.fonts import addMapping

def init():
    global doc, elements, sample_style_sheet, style, blank

    doc = SimpleDocTemplate("create_report.pdf", pagesize=letter, encoding="UTF-8", topMargin=15, bottomMargin=32,)

    elements = []
    sample_style_sheet = getSampleStyleSheet()

    # Register Font for Unicode-utf8
    pdfmetrics.registerFont(TTFont('arial', 'arial.ttf')) # devaNagri is a folder located in **/usr/share/fonts/truetype** and 'NotoSerifDevanagari.ttf' file you just download it from https://www.google.com/get/noto/#sans-deva and move to devaNagri folder.
    pdfmetrics.registerFont(TTFont('arialbd', 'arialbd.ttf'))


    addMapping('arial', 0, 0, 'arial') #devnagri is a folder name and NotoSerifDevanagari is file name 
    addMapping('arialbd', 0, 0, 'arialbd')
    

    style = getSampleStyleSheet()   
    style.add(ParagraphStyle(name="ParagraphTitle", alignment=TA_CENTER, fontName="arial",parent=style['Heading3'], textColor=colors.red, fontSize=8)) # after mapping fontName define your folder name.   
    style.add(ParagraphStyle(name="ParagraphLabel", alignment=TA_CENTER, fontName="arial",parent=style['BodyText'], textColor=colors.black))
    style.add(ParagraphStyle(name="ParagraphNormal", alignment=TA_LEFT, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=4))
    style.add(ParagraphStyle(name="ParagraphNormalRight", alignment=TA_RIGHT, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=4))
    style.add(ParagraphStyle(name="ParagraphNormalCenter", alignment=TA_CENTER, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=4))
    style.add(ParagraphStyle(name="Paragraph4Normal", alignment=TA_LEFT, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=4))

    blank = Paragraph("", style["ParagraphLabel"])

def init_CSV(data,data_monthly):
    global doc, elements, sample_style_sheet, style, blank

    print("Creating Payslip for User: " + data['Employee_Name'])

    doc = SimpleDocTemplate(data['Employee_Code'] + "_" + data['Employee_Name'] + "_" + data_monthly['Monthly_code']+".pdf", encrypt=pdfencrypt.StandardEncryption( data['Employee_Pass'], canPrint=0), pagesize=letter, encoding="UTF-8", topMargin=15, bottomMargin=32,)

    elements = []
    sample_style_sheet = getSampleStyleSheet()

    # Register Font for Unicode-utf8
    pdfmetrics.registerFont(TTFont('arial', 'arial.ttf')) # devaNagri is a folder located in **/usr/share/fonts/truetype** and 'NotoSerifDevanagari.ttf' file you just download it from https://www.google.com/get/noto/#sans-deva and move to devaNagri folder.
    pdfmetrics.registerFont(TTFont('arialbd', 'arialbd.ttf'))

    addMapping('arial', 0, 0, 'arial') #devnagri is a folder name and NotoSerifDevanagari is file name 
    addMapping('arialbd', 0, 0, 'arialbd')

    style = getSampleStyleSheet()   
    style.add(ParagraphStyle(name="ParagraphTitle", alignment=TA_CENTER, fontName="arialbd",parent=style['Heading3'], textColor=colors.red,spaceAfter=0,spaceBefore=0,leading=9)) # after mapping fontName define your folder name.   
    style.add(ParagraphStyle(name="ParagraphLabel", alignment=TA_CENTER, fontName="arial",parent=style['BodyText']))
    style.add(ParagraphStyle(name="ParagraphNormal", alignment=TA_LEFT, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    style.add(ParagraphStyle(name="ParagraphNormalRight", alignment=TA_RIGHT, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    style.add(ParagraphStyle(name="ParagraphNormalLeft", alignment=TA_LEFT, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    style.add(ParagraphStyle(name="ParagraphNormalCenter", alignment=TA_CENTER, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    
    style.add(ParagraphStyle(name="ParagraphHeaderRight", alignment=TA_RIGHT, fontName="arial",parent=style['Normal'], textColor=colors.white, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    style.add(ParagraphStyle(name="ParagraphHeaderLeft", alignment=TA_LEFT, fontName="arial",parent=style['Normal'], textColor=colors.white, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    style.add(ParagraphStyle(name="ParagraphHeaderCenter", alignment=TA_CENTER, fontName="arial",parent=style['Normal'], textColor=colors.white, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))

    style.add(ParagraphStyle(name="ParagraphHeaderBigRight", alignment=TA_RIGHT, fontName="arialbd",parent=style['Heading3'], textColor=colors.white, fontSize=5,spaceAfter=0,spaceBefore=0,leading=7))
    style.add(ParagraphStyle(name="ParagraphHeaderBigLeft", alignment=TA_LEFT, fontName="arialbd",parent=style['Heading3'], textColor=colors.white, fontSize=5,spaceAfter=0,spaceBefore=0,leading=7))
    style.add(ParagraphStyle(name="ParagraphHeaderBigCenter", alignment=TA_CENTER, fontName="arialbd",parent=style['Heading3'], textColor=colors.white, fontSize=5,spaceAfter=0,spaceBefore=0,leading=7))

    style.add(ParagraphStyle(name="ParagraphHeaderBoldRight", alignment=TA_RIGHT, fontName="arialbd",parent=style['Normal'], textColor=colors.white, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    style.add(ParagraphStyle(name="ParagraphHeaderBoldLeft", alignment=TA_LEFT, fontName="arialbd",parent=style['Normal'], textColor=colors.white, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    style.add(ParagraphStyle(name="ParagraphHeaderBoldCenter", alignment=TA_CENTER, fontName="arialbd",parent=style['Normal'], textColor=colors.white, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    style.add(ParagraphStyle(name="ParagraphBigBoldRight", alignment=TA_RIGHT, fontName="arialbd",parent=style['Normal'], textColor=colors.black, fontSize=7,spaceAfter=0,spaceBefore=0,leading=8))
    style.add(ParagraphStyle(name="ParagraphBoldCenter", alignment=TA_CENTER, fontName="arialbd",parent=style['Normal'], textColor=colors.black, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    
    style.add(ParagraphStyle(name="Paragraph4Normal", alignment=TA_LEFT, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))
    style.add(ParagraphStyle(name="ParagraphNormal9", alignment=TA_CENTER, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=4,spaceAfter=0,spaceBefore=0,leading=5))

    blank = Paragraph("", style["ParagraphLabel"])

def create_Title():
    # Part 1
    paragraphTitle = Paragraph('PAYSLIP', style["ParagraphTitle"])
    paragraphLabel = Paragraph("PHIẾU LƯƠNG", style["ParagraphLabel"])

    # image
    I = Image('logo_sparx.png')
    I.drawHeight = 1.25*50*I.drawHeight / I.drawWidth
    I.drawWidth = 70

    table1_data= [[I, paragraphTitle,''],
                  ['', paragraphLabel,'']]
    table1=Table(table1_data, [199,199,199])
    table1_style = [('ALIGN',(0,0),(0,1),'LEFT'),('ALIGN',(1,0),(1,1),'CENTER'),('SPAN', (0,0), (0,1))]
    table1.setStyle(TableStyle(table1_style))
    elements.append(table1)
    elements.append(blank)

    # elements.append()
    # elements.append(paragraphTitle)
    # elements.append(paragraphLabel)
    # elements.append(blank)

def create_Table1():
    # Part 2
    
    blank = Paragraph("", style["ParagraphLabel"])
    P200_1 = Paragraph("Payment for", style["ParagraphNormal"])
    P200_2 = Paragraph("Kỳ tính lương tháng", style["ParagraphNormal"])
    P210_1 = Paragraph("01/2021", style["ParagraphNormalRight"])
    P202_1 = Paragraph("Exchange rate", style["ParagraphNormal"])
    P202_2 = Paragraph("Tỉ giá", style["ParagraphNormal"])
    P212_1 = Paragraph("24,494,00", style["ParagraphNormalRight"])
    P222_1 = Paragraph("VND/USD (Vietcombank exchange rate by 25th of the month)", style["ParagraphNormal"])
    P222_2 = Paragraph("VND/USD (tỉ giá ngân hàng Vietcombank vào ngày 25 hằng tháng)", style["ParagraphNormal"])
    

    table1_data= [[[P200_1,P200_2], P210_1,'',''],
                  ['', '','',''],
                  [[P202_1,P202_2], P212_1,[P222_1,P222_2],''],
                  ['', '','',''],]
    table1=Table(table1_data, [147,147,147,147],15)
    table1_style = [('FONTSIZE', (0,0), (-1,-1), 4),
                    ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                    ('INNERGRID', (1,0), (2,3), 0.25, colors.black),
                    ('BACKGROUND', (0, 0), (-1, -1), colors.aliceblue),
                    ('SPAN', (0,0), (0,1)),('SPAN', (0,2), (0,3)),
                    ('SPAN', (1,0), (1,1)),('SPAN', (2,0), (2,1)),
                    ('SPAN', (1,2), (1,3)),('SPAN', (2,2), (3,3)),
                    ('ALIGN',(1,0),(1,1),'RIGHT'),
                    ('ALIGN',(1,2),(1,3),'RIGHT'),
                    ('LINEABOVE',(0,2),(3,2),1,colors.gray),
                    ]
    table1.setStyle(TableStyle(table1_style))
    elements.append(table1)

def create_Table1_CSV(data,data_monthly):
    # Part 2
    
    blank = Paragraph("", style["ParagraphLabel"])
    P200_1 = Paragraph("Payment for", style["ParagraphNormal"])
    P200_2 = Paragraph("Kỳ tính lương tháng", style["ParagraphNormal"])
    P210_1 = Paragraph(data_monthly['Payment_for'], style["ParagraphBigBoldRight"])
    P202_1 = Paragraph("Exchange rate", style["ParagraphNormal"])
    P202_2 = Paragraph("Tỉ giá", style["ParagraphNormal"])
    P212_1 = Paragraph(data_monthly['Exchange_rate'], style["ParagraphNormalRight"])
    P222_1 = Paragraph("VND/USD (Vietcombank exchange rate by 25th of the month)", style["ParagraphNormal"])
    P222_2 = Paragraph("VND/USD (tỉ giá ngân hàng Vietcombank vào ngày 25 hằng tháng)", style["ParagraphNormal"])
    

    table1_data= [[[P200_1,P200_2], P210_1,'',''],
                  ['', '','',''],
                  [[P202_1,P202_2], P212_1,[P222_1,P222_2],''],
                  ['', '','',''],]
    table1=Table(table1_data, [147,147,147,147],7)
    table1_style = [('FONTSIZE', (0,0), (-1,-1), 4),
                    ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                    ('INNERGRID', (1,0), (2,3), 0.25, colors.black),
                    ('BACKGROUND', (0, 0), (-1, -1), colors.lightgrey),
                    ('SPAN', (0,0), (0,1)),('SPAN', (0,2), (0,3)),
                    ('SPAN', (1,0), (1,1)),('SPAN', (2,0), (2,1)),
                    ('SPAN', (1,2), (1,3)),('SPAN', (2,2), (3,3)),
                    ('ALIGN',(1,0),(1,1),'RIGHT'),
                    ('ALIGN',(1,2),(1,3),'RIGHT'),
                    ('LINEABOVE',(0,2),(3,2),1,colors.gray),
                    ]
    table1.setStyle(TableStyle(table1_style))
    elements.append(table1)

def create_DoubleTables():
    # Part 3
    P300_1 = Paragraph("Employee Code", style["ParagraphNormal"])
    P300_2 = Paragraph("Mã số nhân viên", style["ParagraphNormal"])
    P310_1 = Paragraph("Employee Name", style["ParagraphNormal"])
    P310_2 = Paragraph("Họ và tên", style["ParagraphNormal"])
    P320_1 = Paragraph("Disciplines", style["ParagraphNormal"])
    P320_2 = Paragraph("Bộ phận", style["ParagraphNormal"])
    P330_1 = Paragraph("Position", style["ParagraphNormal"])
    P330_2 = Paragraph("Vị trí", style["ParagraphNormal"])
    P340_1 = Paragraph("Basic salary", style["ParagraphNormal"])
    P340_2 = Paragraph("Tiền lương cơ bản", style["ParagraphNormal"])
    P350_1 = Paragraph("Incentive bonus", style["ParagraphNormal"])
    P350_2 = Paragraph("Tiền thưởng", style["ParagraphNormal"])
    P360_1 = Paragraph("Total gross", style["ParagraphNormal"])
    P360_2 = Paragraph("Tổng thu nhập trước thuế", style["ParagraphNormal"])
    P370_1 = Paragraph("Number of dependents", style["ParagraphNormal"])
    P370_2 = Paragraph("Số người phụ thuộc", style["ParagraphNormal"])
    
    P300_3 = Paragraph("ACTUAL WORKING DAY", style["ParagraphNormal"])
    P300_4 = Paragraph("Ngày công trong tháng", style["ParagraphNormal"])
    P310_3 = Paragraph("Standard working days in month", style["ParagraphNormal"])
    P310_4 = Paragraph("Ngày công tiêu chuẩn trong tháng", style["ParagraphNormal"])
    P320_3 = Paragraph("Remaining AL(s) & OIL for final payment", style["ParagraphNormal"])
    P320_4 = Paragraph("Phép năm còn lại", style["ParagraphNormal"])
    P330_3 = Paragraph("Total paid days", style["ParagraphNormal"])
    P330_4 = Paragraph("Tổng số ngày tính lương", style["ParagraphNormal"])
    P340_3 = Paragraph("OVERTIME RECORD", style["ParagraphNormal"])
    P340_4 = Paragraph("Số giờ tăng ca", style["ParagraphNormal"])
    P350_3 = Paragraph("Overtime Hour on Weekdays", style["ParagraphNormal"])
    P350_4 = Paragraph("Số giờ tăng ca ngày thường", style["ParagraphNormal"])
    P360_3 = Paragraph("Overtime Hours on Weekend", style["ParagraphNormal"])
    P360_4 = Paragraph("Số giờ tăng ca cuối tuần", style["ParagraphNormal"])
    P370_3 = Paragraph("Overtime Hours on Holidays", style["ParagraphNormal"])
    P370_4 = Paragraph("Số giờ tăng ca ngày lễ", style["ParagraphNormal"])

    part3_data1= [[[P300_1,P300_2], '', '', 'A1', '40'],
                  [[P310_1,P310_2], '', '', 'Nguyen Van A', '41'],
                  [[P320_1,P320_2], '', '', 'QA', '42'],
                  [[P330_1,P330_2], '', '', 'Penetration tester', '43'],
                  [[P340_1,P340_2], '', '', '0.04', '1.000'],
                  [[P350_1,P350_2], '', '', '0.04', '1.000'],
                  [[P360_1,P360_2], '', '', '0.04', '1.000'],
                  [[P370_1,P370_2], '', '', '1.000.00', '47']]

    part3_data2= [[[P300_3,P300_4], 'DAYS'],
                  [[P310_3,P310_4], '23.00'],
                  [[P320_3,P320_4], '1.000.00'],
                  [[P330_3,P330_4], '1.000.00'],
                  [[P340_3,P340_4], 'HOURS'],
                  [[P350_3,P350_4], '2.000.00'],
                  [[P360_3,P360_4], '2.000.00'],
                  [[P370_3,P370_4], '2.000.00']]
    # Size Table, Alignment
    part3_t1=Table(part3_data1,[58,58,58,58,58], 8*[0.4*inch],hAlign='LEFT')

    part3_t1.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.25, colors.black),
                   ('SPAN', (0,0), (2,0)),('SPAN', (3,0), (4,0)),
                   ('SPAN', (0,1), (2,1)),('SPAN', (3,1), (4,1)),
                   ('SPAN', (0,2), (2,2)),('SPAN', (3,2), (4,2)),
                   ('SPAN', (0,3), (2,3)),('SPAN', (3,3), (4,3)),
                   ('SPAN', (0,4), (2,4)),
                   ('SPAN', (0,5), (2,5)),
                   ('SPAN', (0,6), (2,6)),
                   ('SPAN', (0,7), (2,7)),('SPAN', (3,7), (4,7)),
                   ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                   ('BACKGROUND', (0, 0), (2, 7), colors.aliceblue),
                   ('ALIGN',(3,0),(4,3),'CENTER'),
                   ('ALIGN',(3,4),(4,7),'RIGHT'),
                   ('FONTSIZE', (0,0), (-1,-1), 4),
                   ]))
   

   
    part3_t2=Table(part3_data2,[145,145], 8*[0.4*inch],hAlign='RIGHT')
    
    part3_t2.setStyle(
        TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                    ('BOX', (0,0), (-1,-1), 0.25, colors.black),
                    ('BACKGROUND', (0, 0), (1, 0), colors.aliceblue),
                    ('FONTSIZE', (0,0), (-1,-1), 4),
                    ('BACKGROUND', (0, 4), (1, 4), colors.aliceblue),
                    ('ALIGN',(1,0),(1,0),'CENTER'),
                    ('ALIGN',(1,4),(1,4),'CENTER'),
                    ('ALIGN',(1,1),(1,3),'RIGHT'),
                    ('ALIGN',(1,5),(1,7),'RIGHT'),
                   ]))
    
    part3_t = [[part3_t1,part3_t2]]
    part3_temp = Table(part3_t, [300, 300])  
    
    elements.append(blank)
    elements.append(part3_temp)
    elements.append(PageBreak())

def create_DoubleTables_CSV(data):
    # Part 3
    # Col 1, format(P3_Row_Col_Number [1: Left, 2: Right])
    P300_1 = Paragraph("Employee Code", style["ParagraphHeaderBoldLeft"])
    P300_2 = Paragraph("Mã số nhân viên", style["ParagraphHeaderBoldLeft"])
    P310_1 = Paragraph("Employee Name", style["ParagraphHeaderBoldLeft"])
    P310_2 = Paragraph("Họ và tên", style["ParagraphHeaderBoldLeft"])
    P320_1 = Paragraph("Disciplines", style["ParagraphHeaderBoldLeft"])
    P320_2 = Paragraph("Bộ phận", style["ParagraphHeaderBoldLeft"])
    P330_1 = Paragraph("Position", style["ParagraphHeaderBoldLeft"])
    P330_2 = Paragraph("Vị trí", style["ParagraphHeaderBoldLeft"])
    P340_1 = Paragraph("Basic salary", style["ParagraphHeaderBoldLeft"])
    P340_2 = Paragraph("Tiền lương cơ bản", style["ParagraphHeaderBoldLeft"])
    P350_1 = Paragraph("Incentive bonus", style["ParagraphHeaderBoldLeft"])
    P350_2 = Paragraph("Tiền thưởng", style["ParagraphHeaderBoldLeft"])
    P360_1 = Paragraph("Total gross", style["ParagraphHeaderBoldLeft"])
    P360_2 = Paragraph("Tổng thu nhập trước thuế", style["ParagraphHeaderBoldLeft"])
    P370_1 = Paragraph("Number of dependents", style["ParagraphHeaderBoldLeft"])
    P370_2 = Paragraph("Số người phụ thuộc", style["ParagraphHeaderBoldLeft"])
    
    # Col 2, format(P3_Row_Col_Number [1: Left, 2: Right])
    P301_1 = Paragraph(data['Employee_Code'], style["ParagraphNormalCenter"])
    P311_1 = Paragraph(data['Employee_Name'], style["ParagraphBoldCenter"])
    P321_1 = Paragraph(data['Disciplines'], style["ParagraphNormalCenter"])
    P331_1 = Paragraph(data['Position'], style["ParagraphNormalCenter"])
    P341_1 = Paragraph(data['Basic_salary_1'], style["ParagraphNormalCenter"])
    P341_2 = Paragraph(data['Basic_salary_2'], style["ParagraphNormalCenter"])
    P351_1 = Paragraph(data['Incentive_bonus_1'], style["ParagraphNormalCenter"])
    P351_2 = Paragraph(data['Incentive_bonus_2'], style["ParagraphNormalCenter"])
    P361_1 = Paragraph(data['Total_gross_1'], style["ParagraphNormalCenter"])
    P361_2 = Paragraph(data['Total_gross_2'], style["ParagraphNormalCenter"])
    P371_1 = Paragraph(data['Number_of_dependents'], style["ParagraphNormalCenter"])


    # Col 3
    P300_3 = Paragraph("ACTUAL WORKING DAY", style["ParagraphHeaderBoldLeft"])
    P300_4 = Paragraph("Ngày công trong tháng", style["ParagraphHeaderBoldLeft"])
    P310_3 = Paragraph("Standard working days in month", style["ParagraphNormal"])
    P310_4 = Paragraph("Ngày công tiêu chuẩn trong tháng", style["ParagraphNormal"])
    P320_3 = Paragraph("Remaining AL(s) & OIL for final payment", style["ParagraphNormal"])
    P320_4 = Paragraph("Phép năm còn lại", style["ParagraphNormal"])
    P330_3 = Paragraph("Total paid days", style["ParagraphNormal"])
    P330_4 = Paragraph("Tổng số ngày tính lương", style["ParagraphNormal"])
    P340_3 = Paragraph("OVERTIME RECORD", style["ParagraphHeaderBoldLeft"])
    P340_4 = Paragraph("Số giờ tăng ca", style["ParagraphHeaderBoldLeft"])
    P350_3 = Paragraph("Overtime Hour on Weekdays", style["ParagraphNormal"])
    P350_4 = Paragraph("Số giờ tăng ca ngày thường", style["ParagraphNormal"])
    P360_3 = Paragraph("Overtime Hours on Weekend", style["ParagraphNormal"])
    P360_4 = Paragraph("Số giờ tăng ca cuối tuần", style["ParagraphNormal"])
    P370_3 = Paragraph("Overtime Hours on Holidays", style["ParagraphNormal"])
    P370_4 = Paragraph("Số giờ tăng ca ngày lễ", style["ParagraphNormal"])

    # Col 4, format(P3_Row_Col_Number [1: Left, 2: Right])
    
    P301_4 = Paragraph('DAYS', style["ParagraphHeaderBoldCenter"])
    P311_4 = Paragraph(data['working_days'], style["ParagraphNormalRight"])
    P321_4 = Paragraph(data['Remaining_AL'], style["ParagraphNormalRight"])
    P331_4 = Paragraph(data['Total_paid_days'], style["ParagraphNormalRight"])
    P341_4 = Paragraph('HOURS', style["ParagraphHeaderBoldCenter"])
    P351_4 = Paragraph(data['OT_on_Weekdays'], style["ParagraphNormalRight"])
    P361_4 = Paragraph(data['OT_on_Weekend'], style["ParagraphNormalRight"])
    P371_4 = Paragraph(data['OT_on_Holidays'], style["ParagraphNormalRight"])
    # Col 4

    part3_data1= [[[P300_1,P300_2], '', '', P301_1, ''],
                  [[P310_1,P310_2], '', '', P311_1, ''],
                  [[P320_1,P320_2], '', '', P321_1, ''],
                  [[P330_1,P330_2], '', '', P331_1, ''],
                  [[P340_1,P340_2], '', '', P341_1, P341_2],
                  [[P350_1,P350_2], '', '', P351_1, P351_2],
                  [[P360_1,P360_2], '', '', P361_1, P361_2],
                  [[P370_1,P370_2], '', '', P371_1, '']]

    part3_data2= [[[P300_3,P300_4], P301_4],
                  [[P310_3,P310_4], P311_4],
                  [[P320_3,P320_4], P321_4],
                  [[P330_3,P330_4], P331_4],
                  [[P340_3,P340_4], P341_4],
                  [[P350_3,P350_4], P351_4],
                  [[P360_3,P360_4], P361_4],
                  [[P370_3,P370_4], P371_4]]
    # Size Table, Alignment
    part3_t1=Table(part3_data1,[58,58,58,58,58], 14,hAlign='LEFT')

    part3_t1.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.25, colors.black),
                   ('SPAN', (0,0), (2,0)),('SPAN', (3,0), (4,0)),
                   ('SPAN', (0,1), (2,1)),('SPAN', (3,1), (4,1)),
                   ('SPAN', (0,2), (2,2)),('SPAN', (3,2), (4,2)),
                   ('SPAN', (0,3), (2,3)),('SPAN', (3,3), (4,3)),
                   ('SPAN', (0,4), (2,4)),
                   ('SPAN', (0,5), (2,5)),
                   ('SPAN', (0,6), (2,6)),
                   ('SPAN', (0,7), (2,7)),('SPAN', (3,7), (4,7)),
                   ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                   ('BACKGROUND', (0, 0), (2, 7), colors.gray),
                   ('TEXTCOLOR',(0,0),(2,7),colors.white),
                   ('ALIGN',(3,0),(4,3),'CENTER'),
                   ('ALIGN',(3,4),(4,7),'RIGHT'),
                   ('FONTSIZE', (0,0), (-1,-1), 4),
                   ]))
   

   
    part3_t2=Table(part3_data2,[145,145], 14,hAlign='RIGHT')
    
    part3_t2.setStyle(
        TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                    ('BOX', (0,0), (-1,-1), 0.25, colors.black),
                    ('FONTSIZE', (0,0), (-1,-1), 4),
                    ('BACKGROUND', (0, 0), (1, 0), colors.gray),
                    ('BACKGROUND', (0, 4), (1, 4), colors.gray),
                    ('TEXTCOLOR',(0,0),(0,-1),colors.whitesmoke),
                    ('ALIGN',(1,0),(1,0),'CENTER'),
                    ('ALIGN',(1,4),(1,4),'CENTER'),
                    ('ALIGN',(1,1),(1,3),'RIGHT'),
                    ('ALIGN',(1,5),(1,7),'RIGHT'),
                   ]))
    
    part3_t = [[part3_t1,part3_t2]]
    part3_temp = Table(part3_t, [300, 300])  
    
    elements.append(blank)
    elements.append(part3_temp)

def create_CurrencyTable():
    # Part 4
    P420_1 = Paragraph("""CURRENCY (USD)""", style["ParagraphHeaderBoldCenter"])
    P420_2 = Paragraph("Đơn vị tiền tệ (USD)", style["ParagraphHeaderBoldCenter"])
    P430_1 = Paragraph("CURRENCY (VND)", style["ParagraphHeaderBoldCenter"])
    P430_2 = Paragraph("Đơn vị tiền tệ (VND)", style["ParagraphHeaderBoldCenter"])
    P450_1 = Paragraph("CURRENCY (USD)", style["ParagraphHeaderBoldCenter"])
    P450_2 = Paragraph("Đơn vị tiền tệ (USD)", style["ParagraphHeaderBoldCenter"])
    P460_1 = Paragraph("CURRENCY (VND)", style["ParagraphHeaderBoldCenter"])
    P460_2 = Paragraph("Đơn vị tiền tệ (VND)", style["ParagraphHeaderBoldCenter"])

    P411_1 = Paragraph("Actual basic salary", style["Paragraph4Normal"])
    P411_2 = Paragraph("Tiền lương cơ bản thực tế trong tháng", style["Paragraph4Normal"])
    P412_1 = Paragraph("Actual incentive bonus", style["Paragraph4Normal"])
    P412_2 = Paragraph("Tiền thưởng thực tế trong tháng", style["Paragraph4Normal"])
    P413_1 = Paragraph("% Extra Incentive Bonus compared to actual gross salary", style["Paragraph4Normal"])
    P413_2 = Paragraph("% Tiền thưởng vượt mức so với tổng thu nhập thực tế", style["Paragraph4Normal"])
    P414_1 = Paragraph("Quarterly Bonus", style["Paragraph4Normal"])
    P414_2 = Paragraph("Tiền thưởng quý", style["Paragraph4Normal"])

    P415_1 = Paragraph("Annual Bonus", style["Paragraph4Normal"])
    P415_2 = Paragraph("Tiền thưởng năm", style["Paragraph4Normal"])
    P416_1 = Paragraph("Referral Bonus", style["Paragraph4Normal"])
    P416_2 = Paragraph("Tiền thưởng giới thiệu ứng viên", style["Paragraph4Normal"])
    P417_1 = Paragraph("Other Bonus", style["Paragraph4Normal"])
    P417_2 = Paragraph("Tiền thưởng khác", style["Paragraph4Normal"])
    P418_1 = Paragraph("Taxable OT payment", style["Paragraph4Normal"])
    P418_2 = Paragraph("Tiền lương tăng ca chịu thuế", style["Paragraph4Normal"])
    P419_1 = Paragraph("Non-taxable OT payment", style["Paragraph4Normal"])
    P419_2 = Paragraph("Tiền lương tăng ca không chịu thuế", style["Paragraph4Normal"])
    P4110_1 = Paragraph("Meal allowance for Overtime (taxable income)", style["Paragraph4Normal"])
    P4110_2 = Paragraph("Tiền ăn tăng ca (chịu thuế)", style["Paragraph4Normal"])
    P4111_1 = Paragraph("Meal allowance for Overtime (non taxable income)", style["Paragraph4Normal"])
    P4111_2 = Paragraph("Tiền ăn tăng ca (không chịu thuế)", style["Paragraph4Normal"])
    P4112_1 = Paragraph("Actual Responsibility Allowance", style["Paragraph4Normal"])
    P4112_2 = Paragraph("Phụ cấp trách nhiệm", style["Paragraph4Normal"])
    P4113_1 = Paragraph("Parking Allowance", style["Paragraph4Normal"])
    P4113_2 = Paragraph("Phụ cấp gửi xe", style["Paragraph4Normal"])
    P4114_1 = Paragraph("Other taxable income", style["Paragraph4Normal"])
    P4114_2 = Paragraph("Các khoản thu nhập chịu thuế khác", style["Paragraph4Normal"])
    P4115_1 = Paragraph("Benefit-in-kind (for PIT only)", style["Paragraph4Normal"])
    P4115_2 = Paragraph("Các khoản thu nhập chịu thuế ngoài bảng lương", style["Paragraph4Normal"])
    P4116_1 = Paragraph("NET PAY", style["ParagraphNormalCenter"])
    P4116_2 = Paragraph("Thực nhận", style["ParagraphNormalCenter"])
    P4116_3 = Paragraph("[25] = [1]+[2]+[3]+[4]+[5]+[6]+[7]+[8]+[9]+[10]+[11]+[12]+[13]-[15]-[16]-[17]-[18]-[19]-[20]-[21]+[22]+[23]", style["ParagraphNormalCenter"])
    P4118_1 = Paragraph("PAYMENT OUT OF PAYROLL PERIOD", style["ParagraphNormalCenter"])
    P4118_2 = Paragraph("Thanh toán ngoài kỳ lương", style["ParagraphNormalCenter"])
    P4120_1 = Paragraph("BANK TRANSFER", style["ParagraphNormalCenter"])
    P4120_2 = Paragraph("Chuyển khoản", style["ParagraphNormalCenter"])
    P4120_3 = Paragraph("[27] = [25]-[26]", style["ParagraphNormalCenter"])





    P441_1 = Paragraph("Mandatory insurance deduction", style["Paragraph4Normal"])
    P441_2 = Paragraph("BHXH, BHYT, BHTN", style["Paragraph4Normal"])
    P442_1 = Paragraph("Personal income tax (PIT)", style["Paragraph4Normal"])
    P442_2 = Paragraph("Thuế thu nhập cá nhân", style["Paragraph4Normal"])
    P443_1 = Paragraph("Finalised PIT", style["Paragraph4Normal"])
    P443_2 = Paragraph("Quyết toán thuế", style["Paragraph4Normal"])
    P444_1 = Paragraph("Union fee", style["Paragraph4Normal"])
    P444_2 = Paragraph("Công đoàn phí", style["Paragraph4Normal"])

    P445_1 = Paragraph("Deduction before Tax", style["Paragraph4Normal"])
    P445_2 = Paragraph("Các khoản trừ trước thuế", style["Paragraph4Normal"])
    P446_1 = Paragraph("Deduction after Tax", style["Paragraph4Normal"])
    P446_2 = Paragraph("Các khoản trừ sau thuế", style["Paragraph4Normal"])
    P447_1 = Paragraph("Deduction of Cash Advance", style["Paragraph4Normal"])
    P447_2 = Paragraph("Trừ các khoản tạm ứng", style["Paragraph4Normal"])
    P448_1 = Paragraph("Non-taxable Income", style["Paragraph4Normal"])
    P448_2 = Paragraph("Thu nhập không chịu thuế khác", style["Paragraph4Normal"])
    P449_1 = Paragraph("Social Fund", style["Paragraph4Normal"])
    P449_2 = Paragraph("Trợ cấp BHXH", style["Paragraph4Normal"])
    P4410_1 = Paragraph("Số tiền thưởng LTI và lãi tính đến quý hiện tại", style["Paragraph4Normal"])
    P4410_2 = Paragraph("LTI and interest up to this quarter", style["Paragraph4Normal"])
    P4411_1 = Paragraph("Remarks for Incomes/Deductions in payroll", style["Paragraph4Normal"])
    P4411_2 = Paragraph("Ghi chú cho các khoản thu nhập/khoản trừ lương", style["Paragraph4Normal"])
    P4511_1 = Paragraph("[6] Other Bonus VND200,000", style["Paragraph4Normal"])
    

    part4_data= [['', '', [P420_1,P420_2], [P430_1,P430_2], '','', [P450_1,P450_2], [P460_1,P460_2]],
                 ['[1]', [P411_1,P411_2], '0.04', '1.000.00', '[15]', [P441_1,P441_2], '0.04','1.000'],
                 ['[2]', [P412_1,P412_2], '0.04', '1.000.00', '[16]', [P442_1,P442_2], '0.08','2.000'],
                 ['', [P413_1,P413_2], '0.00%', '', '[17]', [P443_1,P443_2], '0.04','1.000'],
                 ['[3]', [P414_1,P414_2], '0.04', '1.000.00', '[18]', [P444_1,P444_2], '0.04','1.000'],
                 ['[4]', [P415_1,P415_2], '0.04', '1.000.00', '[19]', [P445_1,P445_2], '0.04','1.000'],
                 ['[5]', [P416_1,P416_2], '0.04', '1.000.00', '[20]', [P446_1,P446_2], '0.04','1.000'],
                 ['[6]', [P417_1,P417_2], '0.04', '1.000.00', '[21]', [P447_1,P447_2], '0.04','1.000'],
                 ['[7]', [P418_1,P418_2], '0.04', '1.000.00', '[22]', [P448_1,P448_2], '0.04','1.000'],
                 ['[8]', [P419_1,P419_2], '0.04', '1.000.00', '[23]', [P449_1,P449_2], '0.04','1.000'],
                 ['[9]', [P4110_1,P4110_2], '0.04', '1.000.00', '[24]', [P4410_1,P4410_2], '0.04','1.000'],
                 ['[10]', [P4111_1,P4111_2], '0.04', '1.000.00', '', [P4411_1,P4411_2], P4511_1 ,'1.000'],
                 ['[11]', [P4112_1,P4112_2], '0.04', '1.000.00', '', '', '0.04','1.000'],
                 ['[12]', [P4113_1,P4113_2], '0.04', '1.000.00', '', '', '0.04','1.000'],
                 ['[13]', [P4114_1,P4114_2], '0.04', '1.000.00', '', '', '0.04','1.000'],
                 ['[14]', [P4115_1,P4115_2], '0.04', '1.000.00', '', '', '0.04','1.000'],
                 ['[25]', [P4116_1,P4116_2,P4116_3], '0.04', '1.000.00', '', '', '0.04','1.000'],
                 ['', '', '', '', '', '', '',''],
                 ['[26]', [P4118_1,P4118_2], '0.04', '1.000.00', '[18]', '', '0.04','1.000'],
                 ['', '', '', '', '', '', '',''],
                 ['[27]', [P4120_1,P4120_2,P4120_3], '0.04', '1.000.00', '[18]', '', '-','1.000'],
                 ['', '', '', '', '', '', '',''],]
    
    part4_t1=Table(part4_data, colWidths=[27,150,62,58,27,123,70,70], rowHeights=30)

    part4_t1.setStyle(
        TableStyle([('BOX', (0,0), (-1,-1), 0.25, colors.black),
                    ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                    ('BACKGROUND', (0, 0), (21, 0), colors.aliceblue),
                    ('BACKGROUND', (0, 16), (7, 17), colors.gray),
                    ('BACKGROUND', (0, 18), (7, 19), colors.lightgrey),
                    ('BACKGROUND', (0, 20), (7, 21), colors.lightgrey),
                    ('SPAN', (2,3), (3,3)),('ALIGN',(2,3), (3,3),'RIGHT'),
                    ('SPAN', (0,16), (0,17)),('SPAN', (1,16), (5,17)),('SPAN', (6,16), (6,17)),('SPAN', (7,16), (7,17)),
                    ('SPAN', (0,18), (0,19)),('SPAN', (1,18), (5,19)),('SPAN', (6,18), (6,19)),('SPAN', (7,18), (7,19)),
                    ('SPAN', (0,20), (0,21)),('SPAN', (1,20), (5,21)),('SPAN', (6,20), (6,21)),('SPAN', (7,20), (7,21)),
                    ('SPAN', (5,11), (5,15)),('SPAN', (6,11), (7,15)),
                    ('FONTSIZE', (0,0), (-1,-1), 4),
                    ]))
    elements.append(part4_t1)

def create_CurrencyTable_CSV(data):
    # Part 4
    P420_1 = Paragraph("""CURRENCY (USD)""", style["ParagraphHeaderBoldCenter"])
    P420_2 = Paragraph("Đơn vị tiền tệ (USD)", style["ParagraphHeaderBoldCenter"])
    P430_1 = Paragraph("CURRENCY (VND)", style["ParagraphHeaderBoldCenter"])
    P430_2 = Paragraph("Đơn vị tiền tệ (VND)", style["ParagraphHeaderBoldCenter"])
    P450_1 = Paragraph("CURRENCY (USD)", style["ParagraphHeaderBoldCenter"])
    P450_2 = Paragraph("Đơn vị tiền tệ (USD)", style["ParagraphHeaderBoldCenter"])
    P460_1 = Paragraph("CURRENCY (VND)", style["ParagraphHeaderBoldCenter"])
    P460_2 = Paragraph("Đơn vị tiền tệ (VND)", style["ParagraphHeaderBoldCenter"])

    P411_1 = Paragraph("Actual basic salary", style["Paragraph4Normal"])
    P411_2 = Paragraph("Tiền lương cơ bản thực tế trong tháng", style["Paragraph4Normal"])
    P412_1 = Paragraph("Actual incentive bonus", style["Paragraph4Normal"])
    P412_2 = Paragraph("Tiền thưởng thực tế trong tháng", style["Paragraph4Normal"])
    P413_1 = Paragraph("% Extra Incentive Bonus compared to actual gross salary", style["Paragraph4Normal"])
    P413_2 = Paragraph("% Tiền thưởng vượt mức so với tổng thu nhập thực tế", style["Paragraph4Normal"])
    P414_1 = Paragraph("Quarterly Bonus", style["Paragraph4Normal"])
    P414_2 = Paragraph("Tiền thưởng quý", style["Paragraph4Normal"])

    P415_1 = Paragraph("Annual Bonus", style["Paragraph4Normal"])
    P415_2 = Paragraph("Tiền thưởng năm", style["Paragraph4Normal"])
    P416_1 = Paragraph("Referral Bonus", style["Paragraph4Normal"])
    P416_2 = Paragraph("Tiền thưởng giới thiệu ứng viên", style["Paragraph4Normal"])
    P417_1 = Paragraph("Other Bonus", style["Paragraph4Normal"])
    P417_2 = Paragraph("Tiền thưởng khác", style["Paragraph4Normal"])
    P418_1 = Paragraph("Taxable OT payment", style["Paragraph4Normal"])
    P418_2 = Paragraph("Tiền lương tăng ca chịu thuế", style["Paragraph4Normal"])
    P419_1 = Paragraph("Non-taxable OT payment", style["Paragraph4Normal"])
    P419_2 = Paragraph("Tiền lương tăng ca không chịu thuế", style["Paragraph4Normal"])
    P4110_1 = Paragraph("Meal allowance for Overtime (taxable income)", style["Paragraph4Normal"])
    P4110_2 = Paragraph("Tiền ăn tăng ca (chịu thuế)", style["Paragraph4Normal"])
    P4111_1 = Paragraph("Meal allowance for Overtime (non taxable income)", style["Paragraph4Normal"])
    P4111_2 = Paragraph("Tiền ăn tăng ca (không chịu thuế)", style["Paragraph4Normal"])
    P4112_1 = Paragraph("Actual Responsibility Allowance", style["Paragraph4Normal"])
    P4112_2 = Paragraph("Phụ cấp trách nhiệm", style["Paragraph4Normal"])
    P4113_1 = Paragraph("Parking Allowance", style["Paragraph4Normal"])
    P4113_2 = Paragraph("Phụ cấp gửi xe", style["Paragraph4Normal"])
    P4114_1 = Paragraph("Other taxable income", style["Paragraph4Normal"])
    P4114_2 = Paragraph("Các khoản thu nhập chịu thuế khác", style["Paragraph4Normal"])
    P4115_1 = Paragraph("Benefit-in-kind (for PIT only)", style["Paragraph4Normal"])
    P4115_2 = Paragraph("Các khoản thu nhập chịu thuế ngoài bảng lương", style["Paragraph4Normal"])
    P4116_1 = Paragraph("NET PAY", style["ParagraphHeaderBigCenter"])
    P4116_2 = Paragraph("Thực nhận", style["ParagraphHeaderBigCenter"])
    P4116_3 = Paragraph("[25] = [1]+[2]+[3]+[4]+[5]+[6]+[7]+[8]+[9]+[10]+[11]+[12]+[13]-[15]-[16]-[17]-[18]-[19]-[20]-[21]+[22]+[23]", style["ParagraphHeaderCenter"])
    P4118_1 = Paragraph("PAYMENT OUT OF PAYROLL PERIOD", style["ParagraphHeaderBigCenter"])
    P4118_2 = Paragraph("Thanh toán ngoài kỳ lương", style["ParagraphHeaderBigCenter"])
    P4120_1 = Paragraph("BANK TRANSFER", style["ParagraphHeaderBigCenter"])
    P4120_2 = Paragraph("Chuyển khoản", style["ParagraphHeaderBigCenter"])
    P4120_3 = Paragraph("[27] = [25]-[26]", style["ParagraphHeaderCenter"])


    P441_1 = Paragraph("Mandatory insurance deduction", style["Paragraph4Normal"])
    P441_2 = Paragraph("BHXH, BHYT, BHTN", style["Paragraph4Normal"])
    P442_1 = Paragraph("Personal income tax (PIT)", style["Paragraph4Normal"])
    P442_2 = Paragraph("Thuế thu nhập cá nhân", style["Paragraph4Normal"])
    P443_1 = Paragraph("Finalised PIT", style["Paragraph4Normal"])
    P443_2 = Paragraph("Quyết toán thuế", style["Paragraph4Normal"])
    P444_1 = Paragraph("Union fee", style["Paragraph4Normal"])
    P444_2 = Paragraph("Công đoàn phí", style["Paragraph4Normal"])

    P445_1 = Paragraph("Deduction before Tax", style["Paragraph4Normal"])
    P445_2 = Paragraph("Các khoản trừ trước thuế", style["Paragraph4Normal"])
    P446_1 = Paragraph("Deduction after Tax", style["Paragraph4Normal"])
    P446_2 = Paragraph("Các khoản trừ sau thuế", style["Paragraph4Normal"])
    P447_1 = Paragraph("Deduction of Cash Advance", style["Paragraph4Normal"])
    P447_2 = Paragraph("Trừ các khoản tạm ứng", style["Paragraph4Normal"])
    P448_1 = Paragraph("Non-taxable Income", style["Paragraph4Normal"])
    P448_2 = Paragraph("Thu nhập không chịu thuế khác", style["Paragraph4Normal"])
    P449_1 = Paragraph("Social Fund", style["Paragraph4Normal"])
    P449_2 = Paragraph("Trợ cấp BHXH", style["Paragraph4Normal"])
    P4410_1 = Paragraph("Số tiền thưởng LTI và lãi tính đến quý hiện tại", style["Paragraph4Normal"])
    P4410_2 = Paragraph("LTI and interest up to this quarter", style["Paragraph4Normal"])
    P4411_1 = Paragraph("Remarks for Incomes/Deductions in payroll", style["Paragraph4Normal"])
    P4411_2 = Paragraph("Ghi chú cho các khoản thu nhập/khoản trừ lương", style["Paragraph4Normal"])
    P4511_1 = Paragraph("[6] Other Bonus VND" + data['Other_Bonus_2'], style["Paragraph4Normal"])

    P4016_1 = Paragraph("<para>[25]</para>", style["ParagraphHeaderBigCenter"])
    P4018_1 = Paragraph("<para>[26]</para>", style["ParagraphHeaderBigCenter"])
    P4020_1 = Paragraph("<para>[27]</para>", style["ParagraphHeaderBigCenter"])
    P4616_1 = Paragraph("<para><font color=yellow>" + data['NET_PAY_1']+ "</font></para>", style["ParagraphHeaderBigCenter"])
    P4616_2 = Paragraph("<para><font color=yellow>" + data['NET_PAY_2']+ "</font></para>", style["ParagraphHeaderBigCenter"])
    P4618_1 = Paragraph(data['PAYMENT_OUT_OF_PAYROLL_PERIOD_1'], style["ParagraphHeaderBigCenter"])
    P4618_2 = Paragraph(data['PAYMENT_OUT_OF_PAYROLL_PERIOD_2'], style["ParagraphHeaderBigCenter"])
    P4620_1 = Paragraph(data['BANK_TRANSFER_1'], style["ParagraphHeaderBigCenter"])
    P4620_2 = Paragraph(data['BANK_TRANSFER_2'], style["ParagraphHeaderBigCenter"])
    

    part4_data= [['', '', [P420_1,P420_2], [P430_1,P430_2], '','', [P450_1,P450_2], [P460_1,P460_2]],
                 ['[1]', [P411_1,P411_2], data['Actual_basic_salary_1'], data['Actual_basic_salary_2'], '[15]', [P441_1,P441_2], data['Mandatory_insurance_1'], data['Mandatory_insurance_2']],
                 ['[2]', [P412_1,P412_2], data['Actual_incentive_bonus_1'], data['Actual_incentive_bonus_2'], '[16]', [P442_1,P442_2], data['PIT_1'], data['PIT_2']],
                 ['', [P413_1,P413_2], data['Extra_Incentive_Bonus'], '', '[17]', [P443_1,P443_2], data['Finalised_PIT_1'], data['Finalised_PIT_2']],
                 ['[3]', [P414_1,P414_2], data['Quarterly_Bonus_1'], data['Quarterly_Bonus_2'], '[18]', [P444_1,P444_2], data['Union_fee_1'], data['Union_fee_2']],
                 ['[4]', [P415_1,P415_2], data['Annual_Bonus_1'], data['Annual_Bonus_2'], '[19]', [P445_1,P445_2], data['Deduction_before_Tax_1'], data['Deduction_before_Tax_2']],
                 ['[5]', [P416_1,P416_2], data['Referral_Bonus_1'], data['Referral_Bonus_2'], '[20]', [P446_1,P446_2], data['Deduction_after_Tax_1'], data['Deduction_after_Tax_2']],
                 ['[6]', [P417_1,P417_2], data['Other_Bonus_1'], data['Other_Bonus_2'], '[21]', [P447_1,P447_2], data['Deduction_of_Cash_Advance_1'], data['Deduction_of_Cash_Advance_2']],
                 ['[7]', [P418_1,P418_2], data['Taxable_OT_payment_1'], data['Taxable_OT_payment_2'], '[22]', [P448_1,P448_2], data['Non_taxable_Income_1'], data['Non_taxable_Income_2']],
                 ['[8]', [P419_1,P419_2], data['Non_taxable_OT_payment_1'], data['Non_taxable_OT_payment_2'], '[23]', [P449_1,P449_2], data['Social_Fund_1'], data['Social_Fund_2']],
                 ['[9]', [P4110_1,P4110_2], data['Overtime_taxable_income_1'], data['Overtime_taxable_income_2'], '[24]', [P4410_1,P4410_2], data['LTI_1'], data['LTI_2']],
                 ['[10]', [P4111_1,P4111_2], data['Overtime_non_taxable_income_1'], data['Overtime_non_taxable_income_2'], '', [P4411_1,P4411_2], P4511_1 ,''],
                 ['[11]', [P4112_1,P4112_2], data['Actual_Responsibility_Allowance_1'], data['Actual_Responsibility_Allowance_2'], '', '', '',''],
                 ['[12]', [P4113_1,P4113_2], data['Parking_Allowance_1'], data['Parking_Allowance_2'], '', '', '',''],
                 ['[13]', [P4114_1,P4114_2], data['Other_taxable_income_1'], data['Other_taxable_income_2'], '', '', '',''],
                 ['[14]', [P4115_1,P4115_2], data['Benefit_in_kind_1'], data['Benefit_in_kind_2'], '', '', '',''],
                 [P4016_1 , [P4116_1,P4116_2,P4116_3], '', '', '', '', P4616_1, P4616_2],
                 ['', '', '', '', '', '', '',''],
                 [P4018_1, [P4118_1,P4118_2], '', '', '[18]', '', P4618_1, P4618_2],
                 ['', '', '', '', '', '', '',''],
                 [P4020_1, [P4120_1,P4120_2,P4120_3], '', '', '[18]', '', P4620_1, P4620_2],
                 ['', '', '', '', '', '', '',''],]
    
    part4_t1=Table(part4_data, colWidths=[27,150,62,58,27,123,70,70], rowHeights=14)

    part4_t1.setStyle(
        TableStyle([('BOX', (0,0), (-1,-1), 0.25, colors.black),
                    ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                    ('BACKGROUND', (0, 0), (21, 0), colors.gray),
                    ('TEXTCOLOR',(0,0),(21,0),colors.whitesmoke),
                    ('BACKGROUND', (0, 16), (7, 17), colors.gray),
                    ('BACKGROUND', (0, 18), (7, 19), colors.lightgrey),
                    ('BACKGROUND', (0, 20), (7, 21), colors.lightgrey),
                    ('SPAN', (2,3), (3,3)),('ALIGN',(2,3), (3,3),'RIGHT'),
                    ('SPAN', (0,16), (0,17)),('SPAN', (1,16), (5,17)),('SPAN', (6,16), (6,17)),('SPAN', (7,16), (7,17)),
                    ('SPAN', (0,18), (0,19)),('SPAN', (1,18), (5,19)),('SPAN', (6,18), (6,19)),('SPAN', (7,18), (7,19)),
                    ('SPAN', (0,20), (0,21)),('SPAN', (1,20), (5,21)),('SPAN', (6,20), (6,21)),('SPAN', (7,20), (7,21)),
                    ('SPAN', (5,11), (5,15)),('SPAN', (6,11), (7,15)),
                    ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                    ('VALIGN',(5,11),(6,11),'TOP'),
                    ('ALIGN',(2,1),(3,2),'RIGHT'),
                    ('ALIGN',(2,4),(3,15),'RIGHT'),
                    ('ALIGN',(6,1),(7,10),'RIGHT'),
                    ('ALIGN',(6,16),(7,21),'RIGHT'),
                    ('FONTSIZE', (0,0), (-1,-1), 4),
                    ]))
    elements.append(blank)
    elements.append(part4_t1)

def writeToDisk():
    # write the document to disk
    doc.build(elements)

def inputCSV():
    global data_list, data_monthly
 
    # Open the CSV file for reading
    with open('payslips.csv', mode='r') as file:
    # Create a CSV reader with DictReader 
        csv_reader = csv.DictReader(file)
 
        # Initialize an empty list to store the dictionaries
        data_list = []
    
        # Iterate through each row in the CSV file
        for row in csv_reader:
            # Append each row (as a dictionary) to the list
            data_list.append(row)
    
    # Read file data monthly
    with open('monthly.csv', mode='r') as file:
    # Create a CSV reader with DictReader 
        csv_reader = csv.DictReader(file)
 
        # Initialize an empty list to store the dictionaries
        data_monthly = []
    
        # Iterate through each row in the CSV file
        for row in csv_reader:
            # Append each row (as a dictionary) to the list
            data_monthly.append(row)
    
    # Print the list of dictionaries

def main():

    inputCSV()
    for data in data_list:
        init_CSV(data,data_monthly[0])
        create_Title()
        create_Table1_CSV(data,data_monthly[0])
        create_DoubleTables_CSV(data)
        create_CurrencyTable_CSV(data)
        writeToDisk()
    
    

if __name__ == '__main__':
    main()
 