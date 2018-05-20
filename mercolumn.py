# -*- coding:gbk -*-
#!/usr/bin/env python -u

# from openpyxl import load_workbook
from xlutils.filter import process,XLRDReader, XLWTWriter
import xlrd
import xlutils.copy
import configuration
import error
import common
import os.path
import info

import sys

def ModifyExcel(fname):
    fname = common.Path(fname)
    if (fname.endswith(".xls")):
        info.DisplayInfo("开始处理文件  "+fname)
        return ModifyXls(fname)
    elif (fname.endswith(".xlsx")):
        info.DisplayInfo("程序不支持xlsx文件，请联系开发人员")
    else:
        error.NotExcelFile()
        common.ExitDueToError()
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

    sheetsName = excelFile.sheet_names()
    # sheetsName = [str(v) for v in sheetsName]
    for sheetIdx,sheetName in enumerate(sheetsName):
        sheet = excelFile.sheet_by_name(sheetName)
        itemNum = 1


        row1 = sheet.row(1)

        start = -1
        end = -1
        fullTxt = ""
        for idx,cell in enumerate(row1):
            if cell.value.strip() == "":
                if start == -1:
                    continue
                else:
                    break
            else:
                fullTxt += cell.value.strip()
                if start == -1:
                    start = idx
                end = idx

        if start == -1:
            info.DisplayInfo("".join[excelFname," 中",sheetName.encode("gbk")," 表单没有合并，请检查该表单"])
            continue
        # if idx < configuration.STARTROW - 1:
        #     continue
        # if row[3].value == "":
        #     continue
        xf_index = sheet.cell_xf_index(1,start)
        saved_style = outStyle[xf_index]
        # if saved_style == None:
        #     print "skdgsdgsdgsdgdsgsdgdsg"*100
        #modifyExcelFile.get_sheet(sheetIdx).write(idx,configuration.ORDERCOL - 1,itemNum,saved_style)
        modifyExcelFile.get_sheet(sheetIdx).write_merge(1,1,0,4,fullTxt,saved_style)
        itemNum += 1

    del excelFile

    try:
        modifyExcelFile.save(excelFname)
    except:
        error.CanNotWrite(excelFname)
        common.ExitDueToError()

def profile(fname):
    if fname.endswith(".xls"):
        ModifyExcel(fname)
    elif (fname.endswith(".xlsx")):
        info.DisplayInfo("程序不支持xlsx文件，请联系开发人员\n程序没有处理文件  "+fname)

def Main():
    error.CreateLog("错误.txt")
    excelName = sys.argv[1]
    excelName = common.Path(excelName)
    if os.path.isfile(excelName):
        ModifyExcel(excelName)
    else:
        common.ProDir(profile,excelName)
    info.DisplayInfo("合并完毕")


if __name__ == "__main__":
    common.SafeRunProgram(Main)
    # excelName = raw_input("拖入目录名或者文件名，如果输入的是目录名，程序自动对该目录及其子目录下的所有excel文件进行处理\n")
    # excelName = excelName.strip('"')
    # # a = os.listdir(excelName)
    # # b = 0
    # # for fn in a:
    # #     b = os.path.join(excelName,fn)
    # # print b
    # #
    # # excelName = raw_input("拖入目录名或者文件名，如果输入的是目录名，程序自动对该目录及其子目录下的所有excel文件进行处理\n")
    # # print excelName
    # # if excelName.startswith('"'):
    # #     excelName = excelName[1:-1]
    # # excelName.replace("\\\\","\\")
    # # print excelName
    # # exit(0)
    # #excelName = '"'+excelName+'"'
    # #print type(excelName)
    # # print type(u"dsg")
    # # print type(r"dsg")
    # # print type("dsg")
    # # print type("草")
    # #excelName.replace("\\","\\\\")
    # #excelName = "F:/myfile/file/鸿图/谭家湾/01 平达大桥目录 - 副本.xls"
    # # f = open(excelName,"w")
    # # f.close()
    # # print excelName
    # # print os.path.exists(excelName)
    # # print os.path.isfile(excelName)
    # # print os.path.isdir(excelName)
    # # exit()
    # if os.path.isfile(excelName):
    #     ModifyExcel(excelName)
    # else:
    #     common.ProDir(profile,excelName)
    # print "\n\n\n"
    # raw_input("合并完毕，输入回车退出程序")
