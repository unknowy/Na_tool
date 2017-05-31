#-*- coding:utf-8 -*-
import os
import xlrd
import zipfile
import time
from xlrd import open_workbook
from xlutils.copy import copy
from docx import Document
from docx.shared import Pt
from docx.shared import Inches
from docx.oxml.ns import qn


def doc_table(table_list):
    document = Document('report.docx')
    table = document.add_table(rows=5, cols=4)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = u'扫描对象'
    hdr_cells[1].text = u'扫描数量'
    hdr_cells[2].text = u'漏洞总数'
    hdr_cells[3].text = u'高危漏洞'
    hdr_cells[4].text = u'中危漏洞'
    hd_cells = table.columns[0].cells

    hd_cells[1].text = u'服务器'
    hd_cells[2].text = u'网络设备'
    hd_cells[3].text = u'终端设备'
    hd_cells[4].text = u'服务器'
    hd_cells[5].text = u'网络设备'
    hd_cells[6].text = u'终端设备'

    for i in xrange(5):
        row_cells = table.rows[i + 1].cells
        row_cells[1].text = str(word[4 * i])
        row_cells[2].text = str(word[4 * i + 1])
        row_cells[3].text = str(word[4 * i + 2])
        row_cells[4].text = str(word[4 * i + 3])
    document.save(u'report.docx')


def doc(word):
    document = Document('report.docx')
    paragraph = document.add_paragraph('\n')
    run = paragraph.add_run(word)
    run.font.name = u'宋体'
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    document.save(u'report.docx')


def doc_table(table_list):
    document = Document('report.docx')
    table = document.add_table(rows=7, cols=5)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = u'扫描对象'
    hdr_cells[1].text = u'扫描数量'
    hdr_cells[2].text = u'漏洞总数'
    hdr_cells[3].text = u'高危漏洞'
    hdr_cells[4].text = u'中危漏洞'
    hd_cells = table.columns[0].cells
    hd_cells[1].text = u'服务器'
    hd_cells[2].text = u'网络设备'
    hd_cells[3].text = u'终端设备'
    hd_cells[4].text = u'服务器'
    hd_cells[5].text = u'网络设备'
    hd_cells[6].text = u'终端设备'

    for i in xrange(6):
        row_cells = table.rows[i + 1].cells
        row_cells[1].text = str(table_list[4 * i])
        row_cells[2].text = str(table_list[4 * i + 1])
        row_cells[3].text = str(table_list[4 * i + 2])
        row_cells[4].text = str(table_list[4 * i + 3])
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


def read_excel_4(list):
    temp1 = (u'终端总数:__' + str(list[8]) + u'__扫描终端个数:__' + str(list[8]) + u'__' + u'\n' +
        u'服务器总数:__' + str(list[0]) + u'__扫描服务器个数:__' + str(list[0]) + u'__' + u'\n' +
        u'数据库总数:__' + str(list[12]) + u'__扫描数据库个数:__' + str(list[12]) + u'__' + u'\n' +
        u'网络设备总数:__' + str(list[4]) + u'__扫描网络设备个数:__' + str(list[4]) + u'__' + u'\n'+
        u'发现漏洞总数:__' + str(list[1] + list[9]) + u'__' + u'\n'+ 
        u'终端漏洞个数：高__' + str(list[10]) + u'__中__' + str(list[11]) + u'__低__' + u'0' + u'__\n' +
        u'服务器漏洞个数：高__' + str(list[2]) + u'__中_' + str(list[3]) + u'__低__' + u'0' + u'__\n' +
        u'数据库漏洞个数：高__' + str(list[13]) + u'__中__' + str(list[14]) + u'__低__' + u'0' + u'__\n'+
        u'网络设备漏洞个数：高__' + str(list[5]) + u'__中__' + str(list[6]) + u'__低__' + u'0' + u'__\n')
    return temp1
def read_excel_3(excel_name):
    output_word=[]
    book=xlrd.open_workbook(excel_name)
    sheet_name=book.sheet_names()[0]  # 获得指定索引的sheet名字
    sheet=book.sheet_by_name(sheet_name)
    station_num=int((sheet.cell_value(2, 1)))
    sheet_name=book.sheet_names()[1]
    sheet=book.sheet_by_name(sheet_name)
    hrisk_sum=int(sum(sheet.col_values(5)[2:30]))
    mrisk_sum=int(sum(sheet.col_values(6)[2:30]))
    risk_sum=hrisk_sum + mrisk_sum
    output_word.append(station_num)
    output_word.append(risk_sum)
    output_word.append(hrisk_sum)
    output_word.append(mrisk_sum)
    return output_word


def read_excel_22(excel_name):
    output_word=[]
    safe_report=[]
    book=xlrd.open_workbook(excel_name)
    sheet_name=book.sheet_names()[0]  # 获得指定索引的sheet名字
    sheet=book.sheet_by_name(sheet_name)
    station_num=int(sheet.cell_value(2, 1))
    sheet_name=book.sheet_names()[1]
    sheet=book.sheet_by_name(sheet_name)
    hrisk_col=int(sum(sheet.col_values(5)[2:50]))
    mrisk_col=int(sum(sheet.col_values(6)[2:50]))
    safe_report.append(u'本月共扫描外网桌面终端' + str(station_num) + u'台，')
    safe_report.append(u'发现' + str(hrisk_col) + u'个高危漏洞、' +
                       str(mrisk_col) + u'个中危漏洞，已整改高危漏洞3个，中危漏洞10个，其余已列入4月整改计划。')
    safe_report.append(u'发现' + str(mrisk_col) + u'个中危漏洞。具体扫描结果见附件。')
    safe_report.append(u'发现' + str(hrisk_col) + u'个高危漏洞。具体扫描结果见附件。')
    safe_report.append(u'无高、中危漏洞。具体扫描结果见附件。')
    if hrisk_col == 0 and mrisk_col == 0:
        output_word.append(safe_report[0])
        output_word.append(safe_report[4])
    if hrisk_col == 1 and mrisk_col == 1:
        output_word.append(safe_report[0])
        output_word.append(safe_report[1])
    if hrisk_col == 1 and mrisk_col == 0:
        output_word.append(safe_report[0])
        output_word.append(safe_report[3])
    if hrisk_col == 0 and mrisk_col == 1:
        output_word.append(safe_report[0])
        output_word.append(safe_report[2])

    return output_word


def read_excel_21(excel_name):

    safe_report=[]
    book=xlrd.open_workbook(excel_name)
    sheet_name=book.sheet_names()[0]  # 获得指定索引的sheet名字
    sheet=book.sheet_by_name(sheet_name)
    station_num=int(sheet.cell_value(2, 1))
    sheet_name=book.sheet_names()[1]
    sheet=book.sheet_by_name(sheet_name)
    hrisk_col=int(sum(sheet.col_values(5)[2:50]))
    mrisk_col=int(sum(sheet.col_values(6)[2:50]))
    # del hrisk_col[0:1]
    # del mrisk_col[0:1]
    safe_report.append(u'本月共扫描内网终端' + str(station_num))
    safe_report.append(u'发现' + str(hrisk_col) + u'个高危、' +
                       str(mrisk_col) + u'个中危漏洞，已整改高危漏洞3个，中危漏洞10个，其余已列入4月整改计划。')
    return safe_report


def read_excel_1(excel_name):
    book=xlrd.open_workbook(excel_name)
    sheet_name=book.sheet_names()[1]  # 获得指定索引的sheet名字
    sheet=book.sheet_by_name(sheet_name)
    col_data=sheet.col_values(1)
    risk_data=[]
    safe_report=[]
    output_word=[]
    try:
        row_portal_server=col_data.index('10.232.80.20')
        portal_server_hrisk=int(sheet.cell_value(row_portal_server, 5))
        portal_server_mrisk=int(sheet.cell_value(row_portal_server, 6))
        risk_data.append(portal_server_hrisk)
        risk_data.append(portal_server_mrisk)
    except ValueError:
        portal_server_hrisk=0
        portal_server_mrisk=0
        risk_data.append(portal_server_hrisk)
        risk_data.append(portal_server_mrisk)
    try:
        row_backup_server=col_data.index('10.232.82.65')
        backup_server_hrisk=int(sheet.cell_value(row_backup_server, 5))
        backup_server_mrisk=int(sheet.cell_value(row_backup_server, 6))
        risk_data.append(backup_server_hrisk)
        risk_data.append(backup_server_mrisk)

    except ValueError:
        backup_server_hrisk=0
        backup_server_mrisk=0
        risk_data.append(backup_server_hrisk)
        risk_data.append(backup_server_mrisk)

    try:
        row_dns_server=col_data.index('10.232.80.1')
        dns_server_hrisk=int(sheet.cell_value(row_dns_server, 5))
        dns_server_mrisk=int(sheet.cell_value(row_dns_server, 6))
        risk_data.append(dns_server_hrisk)
        risk_data.append(dns_server_mrisk)
    except ValueError:
        dns_server_hrisk=0
        dns_server_mrisk=0
        risk_data.append(dns_server_hrisk)
        risk_data.append(dns_server_mrisk)
    try:
        row_desk_server=col_data.index('10.232.80.54')
        desk_server_hrisk=int(sheet.cell_value(row_desk_server, 5))
        desk_server_mrisk=int(sheet.cell_value(row_desk_server, 6))
        risk_data.append(desk_server_hrisk)
        risk_data.append(desk_server_mrisk)
    except ValueError:
        desk_server_hrisk=0
        desk_server_mrisk=0
        risk_data.append(desk_server_hrisk)
        risk_data.append(desk_server_mrisk)
    try:
        row_deskbackup_server=col_data.index('10.232.80.53')
        deskbackup_server_hrisk=int(
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
    document = Document()
    document.save('report.docx')
    for num in range(6):
        un_zip(all_zip_file[num], str(num))
        name_temp = (root + str(num) + '\index.xls')
        target_file.append(name_temp)
# 安全月报第一段
    print ('安全月报第一部分生成中......')
    start_report1 = time.clock()
    safe_report_1 = read_excel_1(target_file[2])
    doc(safe_report_1)
    end_report1 = time.clock()
    print (u'耗时'+str(int((end_report1-start_report1)*1000))+u'毫秒')
# 安全月报第二段
    print ('安全月报第二部分生成中......')
    start_report2 = time.clock()
    safe_report_21 = read_excel_21(target_file[1])
    safe_report_22 = read_excel_22(target_file[3])
    safe_report_2 = safe_report_21 + safe_report_22
    doc(safe_report_2)
    end_report2 = time.clock()
    print (u'耗时'+str(int((end_report2-start_report2)*1000))+u'毫秒')

# 安全月报第三段
    print ('安全月报第三部分生成中......')
    start_report3 = time.clock()
    safe_report_3 = read_excel_3(target_file[2])
    safe_report_3 = safe_report_3 + read_excel_3(target_file[0])
    safe_report_3 = safe_report_3 + read_excel_3(target_file[1])
    safe_report_3 = safe_report_3 + read_excel_3(target_file[5])
    safe_report_3 = safe_report_3 + read_excel_3(target_file[3])
    safe_report_3 = safe_report_3 + read_excel_3(target_file[4])
    doc_table(safe_report_3)
    print safe_report_3
    end_report3 = time.clock()
    print (u'耗时'+str(int((end_report3-start_report3)*1000))+u'毫秒')
    print("安全月报生成完毕")
   
# 督查月报
    print ('督查月报生成中......')
    start_report4 = time.clock()
    safe_report_4=read_excel_4(safe_report_3)
    doc(safe_report_4)
    end_report4 = time.clock()
    print (u'耗时'+str(int((end_report4-start_report4)*1000))+u'毫秒')
    print ('督查月报生成完毕')