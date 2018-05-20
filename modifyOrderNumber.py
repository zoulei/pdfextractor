# -*- coding:gbk -*-

# from openpyxl import load_workbook
from xlutils.filter import process,XLRDReader, XLWTWriter
import xlrd
import xlutils.copy
import configuration
import error
import common
import Constant
import info

import excelcommon

orderStype = None

def GetStyle():
    global orderStype
    if not orderStype:
        orderStype = excelcommon.ReadStyle("dist/format.xls",0,6,0)
    return orderStype

def ModifyExcel(fname):
    fname = common.Path(fname)
    if (fname.endswith(".xls")):
        return ModifyXls(fname)
    elif (fname.endswith(".xlsx")):
        info.DisplayInfo("程序不支持xlsx文件，请联系开发人员")
    else:
        error.NotExcelFile()
        # common.ExitDueToError()
        return None

# def ModifyXlsx(excelFName):
#     excelFile = load_workbook(excelFName,read_only=False)
#     sheetsName = excelFile.get_sheet_names()
#     sheetsName = [str(v) for v in sheetsName]
#
#     for sheetName in sheetsName:
#         itemNum = 1
#         sheet = excelFile.get_sheet_by_name(sheetName)
#         for idx, row in enumerate(sheet.rows):
#             if idx < configuration.STARTROW - 1:
#                 continue
#             if row[3].value == None:
#                 continue
#             pos = "%s%d"%(configuration.GetChar(configuration.ORDERCOL),idx + 1)
#             sheet[pos] = itemNum
#             sheet[pos].style = row[3].style
#             itemNum += 1
#
#     try:
#         excelFile.save(excelFName)
#     except:
#         error.CanNotWrite(excelFName)
#         common.ExitDueToError()

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

    formatFile = xlrd.open_workbook(Constant.FROMATFILE,formatting_info=True)
    formatMo,formatStype = copy2(formatFile)

    # formatFile.sheet_by_index(0)
    xf_index = formatFile.sheet_by_index(0).cell_xf_index(6,0)
    saved_style = formatStype[xf_index]

    xf_index = formatFile.sheet_by_index(0).cell_xf_index(6,3)
    pageStyle = formatStype[xf_index]

    sheetsName = excelFile.sheet_names()
    # sheetsName = [str(v.encode("gbk")) for v in sheetsName]
    for sheetIdx,sheetName in enumerate(sheetsName):
        sheet = excelFile.sheet_by_name(sheetName)
        itemNum = 1

        # print sheetName
        for idx, row in enumerate(sheet.get_rows()):
            # if sheetName == "479":
            #     print idx,row[3].value, "--", itemNum


            if idx < configuration.STARTROW - 1:
                continue
            if row[3].value == "":
                continue

            # if saved_style == None:
            #     print "skdgsdgsdgsdgdsgsdgdsg"*100
            modifyExcelFile.get_sheet(sheetIdx).write(idx,configuration.ORDERCOL - 1,itemNum,saved_style)
            modifyExcelFile.get_sheet(sheetIdx).write(idx,configuration.PAGECOL - 1,row[3].value,pageStyle)
            itemNum += 1

    del excelFile

    try:
        modifyExcelFile.save(excelFname)
    except:
        error.CanNotWrite(excelFname)
        common.ExitDueToError()

if __name__ == "__main__":
    excelName = raw_input()
    ModifyExcel(excelName)