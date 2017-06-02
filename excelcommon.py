# -*- coding:gbk -*-
from xlutils.filter import process,XLRDReader, XLWTWriter
import xlrd

import info
import error

def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb,'unknown.xls'),
        w
        )
    return w.output[0][1], w.style_list

def ModifyXls(excelFname):
    excelFile = xlrd.open_workbook(excelFname,formatting_info=True)
    modifyExcelFile,outStyle = copy2(excelFile)

def ReadStyle(fname, sheetIdx, row, column):
    fname = fname.strip("\"")
    if (fname.endswith(".xls")):
        return ReadStyleXls(fname, sheetIdx, row, column)
    elif (fname.endswith(".xlsx")):
        info.DisplayInfo("程序不支持xlsx文件，请联系开发人员")
    else:
        error.NotExcelFile()
        # common.ExitDueToError()
        return None

def ReadStyleXls(fname, sheetIdx, row, column):
    excelFile = xlrd.open_workbook(fname,formatting_info=True)
    modifyExcelFile,outStyle = copy2(excelFile)

    xf_index = excelFile.get_sheet(sheetIdx).cell_xf_index(row,column)
    saved_style = outStyle[xf_index]
    return saved_style

if __name__ == "__main__":
    fname = raw_input()
    ReadStyle(fname,0,)