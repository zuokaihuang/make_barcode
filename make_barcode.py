from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code39, code93, code128, usps, usps4s, ecc200datamatrix
from reportlab.graphics.shapes import Drawing 
from reportlab.lib.units import mm
from reportlab.graphics import renderPDF

#import reportlab.graphics.barcode
import time
import xlrd



def createBarCodes(c, barcode_value, x, y):
    # barcode_value = "PT620K20080400046"

    # barcode39 = code39.Extended39(barcode_value)
    # barcode39Std = code39.Standard39(barcode_value) #, barHeight=20, stop=1)

    # code93 also has an Extended and MultiWidth version
    # barcode93 = code93.Standard93(barcode_value)

    barcode128 = code128.Code128(barcode_value, barWidth=0.76, barHeight=10, stop=1)
    # the multiwidth barcode appears to be broken 
    # barcode128Multi = code128.MultiWidthBarcode(barcode_value)

    # barcode_usps = usps.POSTNET("50158-9999")

    codes = [barcode128] #, barcode93, barcode128, barcode_usps]



    #for code in codes:
    barcode128.drawOn(c, x, y)
    #    y = y - 15 * mm

    # draw the eanbc8 code
    # barcode_eanbc8 = eanbc.Ean8BarcodeWidget(barcode_value)
    # d = Drawing(50, 10)
    # d.add(barcode_eanbc8)
    # renderPDF.draw(d, c, 15, 555)

    # draw the eanbc13 code
    # barcode_eanbc13 = eanbc.Ean13BarcodeWidget(barcode_value)
    # d = Drawing(50, 10)
    # d.add(barcode_eanbc13)
    # renderPDF.draw(d, c, 15, 465)

    # draw a QR code
    # qr_code = qr.QrCodeWidget('http://blog.csdn.net/webzhuce')
    # bounds = qr_code.getBounds()
    # width = bounds[2] - bounds[0]
    # height = bounds[3] - bounds[1]
    # d = Drawing(45, 45, transform=[45./width,0,0,45./height,0,0])
    # d.add(qr_code)
    # renderPDF.draw(d, c, 15, 405)

now = time.strftime("%Y-%m-%d",time.localtime(time.time()))
fname=now+r"_reportlab.pdf"
c=canvas.Canvas(fname)

excel_path = "sn.xls"
excel = xlrd.open_workbook(excel_path,encoding_override="utf-8")

#返回所有Sheet对象的list
all_sheet = excel.sheets()#Book(工作簿)对象方法
print(all_sheet)

#遍历返回的Sheet对象的list
for each_sheet in all_sheet:
    print(each_sheet)
    print("sheet名称为：",each_sheet.name)#sheet对象方法


#循环遍历每个sheet对象
sheet_name = []
sheet_row  = []
sheet_col  = []
for sheet in all_sheet:
    sheet_name.append(sheet.name)
    #print("该Excel共有{0}个sheet,当前sheet名称为{1},该sheet共有{2}行,{3}列"
    #      .format(len(all_sheet),sheet.name,sheet.nrows,sheet.ncols))
    # print(sheet.col_values(0))#获取指定列的数据
 
    #for each_col in range(sheet.ncols):#依次获得每一列的数据
    #    print("当前为%s列："% each_col )
    #    print(sheet.col_values(each_col ),type(sheet.col_values(each_col )))


    xx = 2 * mm
    yy = 285 * mm
    c.setFontSize(3 * mm)

    next_col = 1

    page_index = 1

    for i in range(sheet.nrows):#依次遍历获得每一行对象
        each_cell_value_row = sheet.row(i)
        sn_=sheet.cell_value(i, 0)
        print(sn_)
        width_ = c.stringWidth(sn_)
        c.drawString(xx, yy + 1 * mm, text = sn_)
        createBarCodes(c, sn_, x=xx + width_ - 5, y=yy)
        yy = yy - 8 * mm
        if yy < 0:
            next_col = next_col - 1
            if next_col==0:
                xx = xx + 105 * mm
            else:
                c.showPage()
                page_index = page_index + 1
                print(page_index)
                c.setFontSize(3 * mm)
                xx = 2 * mm
                next_col = 1
            yy = 285 * mm


#定义要生成的pdf的名称
# c=canvas.Canvas("reportlab.pdf")
#调用函数生成条形码和二维码，并将canvas对象作为参数传递
# createBarCodes(c)
#showPage函数：保存当前页的canvas
c.showPage()
#save函数：保存文件并关闭canvas
c.save()
print("生成成功！！！！！！！！！！！！")
input('按 <Enter> 退出')

