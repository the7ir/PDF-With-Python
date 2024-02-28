
from reportlab.lib import colors
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import letter, inch, landscape
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_JUSTIFY, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.platypus import Paragraph, PageBreak
from reportlab.pdfbase import pdfmetrics      
from reportlab.pdfbase.ttfonts import TTFont   
from reportlab.lib.fonts import addMapping

def init():
    global doc, elements, sample_style_sheet, style, blank

    doc = SimpleDocTemplate("create_report.pdf", pagesize=letter, encoding="UTF-8", topMargin=15, bottomMargin=32,)

    elements = []
    sample_style_sheet = getSampleStyleSheet()

    # Register Font for Unicode-utf8
    pdfmetrics.registerFont(TTFont('arial', 'arial.ttf')) # devaNagri is a folder located in **/usr/share/fonts/truetype** and 'NotoSerifDevanagari.ttf' file you just download it from https://www.google.com/get/noto/#sans-deva and move to devaNagri folder.

    addMapping('arial', 0, 0, 'arial') #devnagri is a folder name and NotoSerifDevanagari is file name 

    style = getSampleStyleSheet()   
    style.add(ParagraphStyle(name="ParagraphTitle", alignment=TA_CENTER, fontName="arial",parent=style['Heading3'], textColor=colors.red)) # after mapping fontName define your folder name.   
    style.add(ParagraphStyle(name="ParagraphLabel", alignment=TA_CENTER, fontName="arial",parent=style['BodyText'], textColor=colors.black))
    style.add(ParagraphStyle(name="ParagraphNormal", alignment=TA_LEFT, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=7))
    style.add(ParagraphStyle(name="ParagraphNormalRight", alignment=TA_RIGHT, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=7))
    style.add(ParagraphStyle(name="Paragraph4Normal", alignment=TA_LEFT, fontName="arial",parent=style['Normal'], textColor=colors.black, fontSize=5))

    blank = Paragraph("", style["ParagraphLabel"])

def create_Title():
    # Part 1
    paragraphTitle = Paragraph('Biển Đẹp Sóng Mơ', style["ParagraphTitle"])
    paragraphLabel = Paragraph("a litle development", style["ParagraphLabel"])

    elements.append(paragraphTitle)
    elements.append(paragraphLabel)
    elements.append(blank)

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
    table1=Table(table1_data, [147,147,147,147],15,)
    table1_style = [('FONTSIZE', (0,0), (-1,-1), 7),
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
                   ('FONTSIZE', (0,0), (-1,-1), 7),
                   ]))
   

   
    part3_t2=Table(part3_data2,[145,145], 8*[0.4*inch],hAlign='RIGHT')
    
    part3_t2.setStyle(
        TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                    ('BOX', (0,0), (-1,-1), 0.25, colors.black),
                    ('BACKGROUND', (0, 0), (1, 0), colors.aliceblue),
                    ('FONTSIZE', (0,0), (-1,-1), 7),
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
    elements.append(blank)

def create_CurrencyTable():
    # Part 4
    P420_1 = Paragraph("""<para align=center spaceb=3><b>CURRENCY <font color=red>(USD)</font></b></para>""", style["Paragraph4Normal"])
    P420_2 = Paragraph("Đơn vị tiền tệ (USD)", style["Paragraph4Normal"])
    P430_1 = Paragraph("CURRENCY (VND)", style["Paragraph4Normal"])
    P430_2 = Paragraph("Đơn vị tiền tệ (VND)", style["Paragraph4Normal"])
    P450_1 = Paragraph("CURRENCY (USD)", style["Paragraph4Normal"])
    P450_2 = Paragraph("Đơn vị tiền tệ (USD)", style["Paragraph4Normal"])
    P460_1 = Paragraph("CURRENCY (VND)", style["Paragraph4Normal"])
    P460_2 = Paragraph("Đơn vị tiền tệ (VND)", style["Paragraph4Normal"])

    P411_1 = Paragraph("Actual basic salary", style["Paragraph4Normal"])
    P411_2 = Paragraph("Tiền lương cơ bản thực tế trong tháng", style["Paragraph4Normal"])
    P412_1 = Paragraph("Actual incentive bonus", style["Paragraph4Normal"])
    P412_2 = Paragraph("Tiền thưởng thực tế trong tháng", style["Paragraph4Normal"])
    P413_1 = Paragraph("% Extra Incentive Bonus compared to actual gross salary", style["Paragraph4Normal"])
    P413_2 = Paragraph("% Tiền thưởng vượt mức so với tổng thu nhập thực tế", style["Paragraph4Normal"])
    P414_1 = Paragraph("Quarterly Bonus", style["Paragraph4Normal"])
    P414_2 = Paragraph("Tiền thưởng quý", style["Paragraph4Normal"])

    P441_1 = Paragraph("Mandatory insurance deduction (9.5%)", style["Paragraph4Normal"])
    P441_2 = Paragraph("BHXH, BHYT, BHTN (9.5%)", style["Paragraph4Normal"])
    P442_1 = Paragraph("Personal income tax (PIT)", style["Paragraph4Normal"])
    P442_2 = Paragraph("Thuế thu nhập cá nhân", style["Paragraph4Normal"])
    P443_1 = Paragraph("2022 Finalised PIT", style["Paragraph4Normal"])
    P443_2 = Paragraph("Quyết toán thuế 2022", style["Paragraph4Normal"])
    P444_1 = Paragraph("Union fee", style["Paragraph4Normal"])
    P444_2 = Paragraph("Công đoàn phí", style["Paragraph4Normal"])

    part4_data= [['', '', [P420_1,P420_2], [P430_1,P430_2], '','', [P450_1,P450_2], [P460_1,P460_2]],
                 ['[1]', [P411_1,P411_2], '0.04', '1.000.00', '[15]', [P441_1,P441_2], '0.04','1.000'],
                 ['[2]', [P412_1,P412_2], '0.04', '1.000.00', '[16]', [P442_1,P442_2], '0.08','2.000'],
                 ['', [P413_1,P413_2], '0.00%', '', '[17]', [P443_1,P443_2], '0.04','1.000'],
                 ['[3]', [P414_1,P414_2], '0.04', '1.000.00', '[18]', [P444_1,P444_2], '0.04','1.000']]
    
    part4_t1=Table(part4_data, colWidths=[27,150,62,58,27,123,70,70], rowHeights=30)

    part4_t1.setStyle(
        TableStyle([('BOX', (0,0), (-1,-1), 0.25, colors.black),
                    ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                    ('BACKGROUND', (0, 0), (21, 0), colors.aliceblue),
                    ('SPAN', (2,3), (3,3)),('ALIGN',(2,3), (3,3),'RIGHT'),
                    ('FONTSIZE', (0,0), (-1,-1), 6),
                    ]))
    elements.append(part4_t1)

def writeToDisk():
    # write the document to disk
    doc.build(elements)

def main():

    init()
    create_Title()
    create_Table1()
    create_DoubleTables()
    create_CurrencyTable()
    writeToDisk()
    

if __name__ == '__main__':
    main()
 