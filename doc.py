# coding=utf-8
from docx import Document
from docx.shared import Pt
from docx.shared import Inches
from docx.oxml.ns import qn
import xlrd
from xlrd import open_workbook
# 打开文档


def excel_read3(list):

    book = xlrd.open_workbook(list)
    sheet_name1 = book.sheet_names()[0]  # 获得指定索引的sheet名字
    sheet = book.sheet_by_name(sheet_name1)
    station_num = (sheet.cell_value(2, 1))
    sheet_name1 = book.sheet_names()[1]
    sheet = book.sheet_by_name(sheet_name1)
    risk_sum1 = sum(sheet.col_values(5)[2:30])
    risk_sum2 = sum(sheet.col_values(6)[2:30])
    print (station_num)
    print (risk_sum1)
    print (risk_sum2)


def doc_table_risk_data(table_list):
    document = Document('report.docx')
    table = document.add_table(rows=10, cols=5)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = u'序号'
    hdr_cells[1].text = u'IP地址'
    hdr_cells[2].text = u'所属部门'
    hdr_cells[3].text = u'使用人'
    hdr_cells[4].text = u'操作系统'
    hdr_cells[5].text = u'风险等级'
    hdr_cells[6].text = u'高危漏洞'
    hdr_cells[7].text = u'中危漏洞'
    hdr_cells[8].text = u'合计'
    hdr_cells[9].text = u'主机风险值'

    for i in xrange(3):
        row_cells = table.rows[i + 1].cells
        row_cells[1].text = str(table_list[3 * i])
        row_cells[2].text = str(table_list[3 * i + 1])
        row_cells[3].text = str(table_list[3 * i + 2])
    document.save(u'report.docx')


if __name__ == "__main__":
    list = [1, 2, 3, 4, 5, 6, 7, 8]
    temp = (u'网络设备总数:__' + str(list[4]) + '__扫描网络设备个数:__' + str(list[4]) + u'__' + u'\n')
    print temp
