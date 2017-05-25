#-*- coding:utf-8 -*-
import os
import xlrd
import zipfile
from xlrd import open_workbook
from xlutils.copy import copy
from docx import Document
from docx.shared import Pt
from docx.shared import Inches
from docx.oxml.ns import qn
def doc():
    document = Document()
    #加入不同等级的标题
    head=document.add_heading(u'信息安全月报添加部分',0)

    # document.add_heading(u'一级标题',1)
    # document.add_heading(u'二级标题',2)
    #添加文本
    paragraph = document.add_paragraph(u'           ')
    # #设置字号
    # run = paragraph.add_run(u'设置字号、')
    # run.font.size = Pt(24)
    # #设置字体
    # run = paragraph.add_run('Set Font,')
    # run.font.name = 'Consolas'
    #设置中文字体
    run = paragraph.add_run(u'本月共扫描内网终端2709台，发现10个高危、31个中危，已整改高危3个，中10个，其余已列入4月整改计划。本月共扫描外网桌面终端41台，无中、高危漏洞。具体扫描结果见附件。、')
    run.font.name=u'方正仿宋_GBK'
    run.font.size = Pt(20)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), u'方正仿宋_GBK')
    #设置斜体
    run = paragraph.add_run(u'斜体、')
    run.italic = True
    #设置粗体
    run = paragraph.add_run(u'粗体').bold = True
    #增加无序列表
    document.add_paragraph(
    u'无序列表元素1', style='List Bullet'
    )
    document.add_paragraph(
    u'无序列表元素2', style='List Bullet'
    )
    #增加有序列表
    document.add_paragraph(
    u'有序列表元素1', style='List Number'
    )
    document.add_paragraph(
    u'有序列表元素2', style='List Number'
    )
    #增加图像（此处用到图像image.bmp，请自行添加脚本所在目录中）

    #document.add_picture('1.png')


    #增加表格
    table = document.add_table(rows=1, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Name'
    hdr_cells[1].text = 'Id'
    hdr_cells[2].text = 'Desc'
    #再增加3行表格元素
    for i in xrange(3):
        row_cells = table.add_row().cells
        row_cells[0].text = 'test'+str(i)
        row_cells[1].text = str(i)
        row_cells[2].text = 'desc'+str(i)
    #增加分页
    document.add_page_break()
    #保存文件
    document.save(u'信息安全月报.doc')
def file_name(file_dir):
    L = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if file.endswith(".zip"):

                file_name = os.path.join(root,file)
                L.append(file_name)
    return L


def un_zip(file_name, name):
    """unzip zip file"""

  #  file_name=raw_input(file_name1)
    zip_file = zipfile.ZipFile(file_name)

    # for names in zip_file.namelist():
    zip_file.extract('index.xls', name)
    zip_file.close()


def write_excel(excel):

    rb = open_workbook(excel)

    # 通过sheet_by_index()获取的sheet没有write()方法
    rs = rb.sheet_by_index(0)

    wb = copy(rb)

    # 通过get_sheet()获取的sheet有write()方法
    ws = wb.get_sheet(0)
    ws.write(1, 1, 'changed!')

    wb.save(excel)
    print("ok")


def read_excel(excel_name):
    book = xlrd.open_workbook(excel_name)

    sheet_name = book.sheet_names()[0]  # 获得指定索引的sheet名字
    # print sheet_name
    sheet = book.sheet_by_name(sheet_name)
    cell_value1 = sheet.cell_value(2, 2)
    print cell_value1


if __name__ == "__main__":
    L = file_name('./')
    a ="./"
  
    doc()
    for num in range(2):
        un_zip(L[num], str(num))

        xlsfile = a+str(num)+'\index.xls'
        print xlsfile
        read_excel(xlsfile)
        write_excel(xlsfile)

    print("ok")
 #   read_excel(index.xls)
