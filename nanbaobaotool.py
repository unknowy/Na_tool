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


def doc(word):
    document = Document('report.docx')
    paragraph = document.add_paragraph('    ')
    run = paragraph.add_run(word)
    run.font.name = u'宋体'
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    document.save(u'report.docx')


def file_name(file_dir):
    L = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if file.endswith(".zip"):

                file_name = os.path.join(root, file)
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


def read_excel_2(excel_name):
    output_word=[]
    book = xlrd.open_workbook(excel_name)
    sheet_name = book.sheet_names()[0]  # 获得指定索引的sheet名字
    sheet = book.sheet_by_name(sheet_name)
    station_num = int(sheet.cell_value(2, 1))
    print station_num
    sheet_name = book.sheet_names()[1]
    sheet = book.sheet_by_name(sheet_name)
    print sheet_name
    hrisk_col = int(sum(sheet.col_values(5)[2:50]))
    mrisk_col = int(sum(sheet.col_values(6)[2:50]))
    # del hrisk_col[0:1]
    # del mrisk_col[0:1]
    print hrisk_col
    print mrisk_col
    


def read_excel(excel_name):
    book = xlrd.open_workbook(excel_name)
    sheet_name = book.sheet_names()[1]  # 获得指定索引的sheet名字
    sheet = book.sheet_by_name(sheet_name)
    col_data = sheet.col_values(1)
    risk_data = []
    safe_report=[]
    output_word=[]
    try:
        row_portal_server = col_data.index('10.232.80.20')
        portal_server_hrisk = int(sheet.cell_value(row_portal_server, 5))
        portal_server_mrisk = int(sheet.cell_value(row_portal_server, 6))
        risk_data.append(portal_server_hrisk)
        risk_data.append(portal_server_mrisk)
    except ValueError:
        portal_server_hrisk = 0
        portal_server_mrisk = 0
        risk_data.append(portal_server_hrisk)
        risk_data.append(portal_server_mrisk)
    try:
        row_backup_server = col_data.index('10.232.82.65')
        backup_server_hrisk = int(sheet.cell_value(row_backup_server, 5))
        backup_server_mrisk = int(sheet.cell_value(row_backup_server, 6))
        risk_data.append(backup_server_hrisk)
        risk_data.append(backup_server_mrisk)

    except ValueError:
        backup_server_hrisk = 0
        backup_server_mrisk = 0
        risk_data.append(backup_server_hrisk)
        risk_data.append(backup_server_mrisk)

    try:
        row_dns_server = col_data.index('10.232.80.1')
        dns_server_hrisk = int(sheet.cell_value(row_dns_server, 5))
        dns_server_mrisk = int(sheet.cell_value(row_dns_server, 6))
        risk_data.append(dns_server_hrisk)
        risk_data.append(dns_server_mrisk)
    except ValueError:
        dns_server_hrisk = 0
        dns_server_mrisk = 0
        risk_data.append(dns_server_hrisk)
        risk_data.append(dns_server_mrisk)
    try:
        row_desk_server = col_data.index('10.232.80.54')
        desk_server_hrisk = int(sheet.cell_value(row_desk_server, 5))
        desk_server_mrisk = int(sheet.cell_value(row_desk_server, 6))
        risk_data.append(desk_server_hrisk)
        risk_data.append(desk_server_mrisk)
    except ValueError:
        desk_server_hrisk = 0
        desk_server_mrisk = 0
        risk_data.append(desk_server_hrisk)
        risk_data.append(desk_server_mrisk)
    try:
        row_deskbackup_server = col_data.index('10.232.80.53')
        deskbackup_server_hrisk = int(
            sheet.cell_value(row_deskbackup_server, 5))
        deskbackup_server_mrisk = int(
            sheet.cell_value(row_deskbackup_server, 6))
        risk_data.append(deskbackup_server_mrisk)
        risk_data.append(deskbackup_server_mrisk)
    except ValueError:
        deskbackup_server_hrisk = 0
        deskbackup_server_mrisk = 0
        risk_data.append(deskbackup_server_mrisk)
        risk_data.append(deskbackup_server_mrisk)

    safe_report.append(u'本月使用绿盟扫描，')
    safe_report.append(u'门户服务器')
    safe_report.append(str(risk_data[0]) + u'个高危漏洞（已备案），')
    safe_report.append(str(risk_data[1]) + u'个中危漏洞（已备案），')
    safe_report.append(u'DNS服务器')
    safe_report.append(str(risk_data[2]) + u'个高危漏洞（已备案），')
    safe_report.append(str(risk_data[3]) + u'个中危漏洞（已备案），')
    safe_report.append(u'备份服务器')
    safe_report.append(str(risk_data[4]) + u'个高危漏洞（已备案），')
    safe_report.append(str(risk_data[5]) + u'个中危漏洞（已备案），')
    safe_report.append(u'桌管服务器')
    safe_report.append(str(risk_data[6]) + u'个高危漏洞（已备案），')
    safe_report.append(str(risk_data[7]) + u'个中危漏洞（已备案），')
    safe_report.append(u'桌管备份服务器')
    safe_report.append(str(risk_data[8]) + u'个高危漏洞（已备案），')
    safe_report.append(str(risk_data[9]) + u'个中危漏洞（已备案），')
    safe_report.append(u'外网主机、服务器、终端均未发现中、高危漏洞。自建系统未发现高、中危漏洞。具体扫描结果见附件。')
    output_word.append(safe_report[0])
    for i in range(5):
        if risk_data[2 * i] != 0 or risk_data[2 * i + 1] != 0:
            output_word.append(safe_report[3 * i + 1])
        if risk_data[2 * i] != 0:
            output_word.append(safe_report[3 * i + 2])
        if risk_data[2 * i + 1] != 0:
            output_word.append(safe_report[3 * i + 3])
    output_word.append(safe_report[16])
    return output_word

if __name__ == "__main__":
    all_zip_file = file_name('./')
    root = "./"
    target_file = []
    safe_report_1 = []
    document = Document()
    document.save('report.docx')
    for num in range(4):
        un_zip(all_zip_file[num], str(num))
        name_temp = (root + str(num) + '\index.xls')
        target_file.append(name_temp)
# 安全月报第一段
    safe_report_1  = read_excel(target_file[2])
    doc(output_word)
# 安全月报第二段
    read_excel_2(target_file[1])

    print("ok")
 #   read_excel(index.xls)
