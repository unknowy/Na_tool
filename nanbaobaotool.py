#-*- coding:utf-8 -*-
import os
import xlrd
import zipfile
from xlrd import open_workbook
from xlutils.copy import copy


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
    for num in range(2):
        un_zip(L[num], str(num))

        xlsfile = a+str(num)+'\index.xls'
        print xlsfile
        read_excel(xlsfile)
        write_excel(xlsfile)

    print("ok")
 #   read_excel(index.xls)
