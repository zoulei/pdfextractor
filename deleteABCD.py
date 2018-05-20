# -*- coding: gbk -*-

# from openpyxl import load_workbook
from xlutils.filter import process,XLRDReader, XLWTWriter
import xlrd
import configuration
import error
import common
import info
import sys

def ModifyExcel(fname):
    fname = common.Path(fname)
    if (fname.endswith(".xls")):
        return ModifyXls(fname)
    elif (fname.endswith(".xlsx")):
        info.DisplayInfo("程序不支持xlsx文件，请联系开发人员")
    else:
        error.NotExcelFile()
        exit()
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
#             pos = "%s%d"%(configuration.GetChar(configuration.PAGECOL),idx + 1)
#             saveStyle = row[3].style
#             sheet[pos] = common.GeneraterPageNumber(str(sheet[pos].value))
#             sheet[pos].style = saveStyle
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

    sheetsName = excelFile.sheet_names()
    sheetsName = [str(v) for v in sheetsName]
    for sheetIdx,sheetName in enumerate(sheetsName):
        sheet = excelFile.sheet_by_name(sheetName)
        itemNum = 1

        for idx, row in enumerate(sheet.get_rows()):
            if idx < configuration.STARTROW - 1:
                continue
            if row[3].value == "":
                continue
            xf_index = sheet.cell_xf_index(idx,configuration.PAGECOL - 1)
            saved_style = outStyle[xf_index]
            pageNum = common.GeneraterPageNumber(str(sheet.cell(idx,configuration.PAGECOL - 1).value) )

            modifyExcelFile.get_sheet(sheetIdx).write(idx,configuration.PAGECOL - 1,pageNum,saved_style)
            itemNum += 1

    del excelFile

    try:
        modifyExcelFile.save(excelFname)
    except:
        error.CanNotWrite(excelFname)
        common.ExitDueToError()

def delABCDMain():
    # info.DisplayInfo( "请输入excel文件路径.")
    error.CreateLog("错误.txt")
    excelFName = sys.argv[1]
    info.DisplayInfo( "开始处理文件")
    ModifyExcel(excelFName)
    info.DisplayInfo( "处理文件结束")

if __name__ == "__main__":
    common.SafeRunProgram(delABCDMain)
