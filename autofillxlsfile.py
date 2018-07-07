# -*- coding: gbk -*-

from xlutils.filter import process,XLRDReader, XLWTWriter
import xlrd
import error
import common
import info
import sys
import traceback

def ModifyExcel(fname,outputExcelFName):
    fname = common.Path(fname)
    outputExcelFName = common.Path(outputExcelFName)
    if (fname.endswith(".xls")):
        return ModifyXls(fname,outputExcelFName)
    elif (fname.endswith(".xlsx")):
        info.DisplayInfo("程序不支持xlsx文件，请联系开发人员")
    else:
        error.NotExcelFile()
        exit()
        return None

def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb,'unknown.xls'),
        w
        )
    return w.output[0][1], w.style_list

def ModifyXls(excelFname,outputexcelFname):
    excelFile = xlrd.open_workbook(excelFname,formatting_info=True)
    sheetsName = excelFile.sheet_names()
    # sheetsName = [str(v) for v in sheetsName]

    # sheetsName = zip(range(len(sheetsName)),sheetsName)
    # sheetsName.sort(key=lambda v:int(v[1]))
    realsheetsname = []
    for sheetname in sheetsName:
        try:
            sheet = excelFile.sheet_by_name(sheetname)
            dananhao = sheet.cell_value(1, 0)
            shunxuidx = dananhao.find(u"－")
            if shunxuidx == -1:
                shunxuidx = dananhao.find("-")
            bianhao = int(dananhao[shunxuidx + 1:].strip())
            realsheetsname.append([bianhao, sheetname])
        except Exception as e:
            info.DisplayInfo("错误发生表单，错误表单名字如下："  )
            info.DisplayInfo(sheetname)
            traceback.print_exc()
            raise e
    realsheetsname.sort(key=lambda v:v[0])
    sheetsName = realsheetsname

    sheetsName = sheetsName[1:] + [sheetsName[0],]

    outputexcelFile = xlrd.open_workbook(outputexcelFname,formatting_info=True)
    modifyExcelFile, outStyle = copy2(outputexcelFile)
    # for idx in xrange(len(outputexcelFile.sheet_names())):
    #     print "sheet name:", outputexcelFile.sheet_names()[idx]
    # sheet = outputexcelFile.sheet_by_name(str(outputexcelFile.sheet_names()[0]))
    sheet = outputexcelFile.sheet_by_name(outputexcelFile.sheet_names()[0])

    xf_index = sheet.cell_xf_index(4, 0)
    stylenumber = outStyle[xf_index]
    xf_index = sheet.cell_xf_index(4, 1)
    styledanganhao = outStyle[xf_index]
    xf_index = sheet.cell_xf_index(4, 2)
    stylecontent = outStyle[xf_index]
    xf_index = sheet.cell_xf_index(4, 6)
    stylezhangci = outStyle[xf_index]

    for idx,sheetName in sheetsName:
        try:
            # sheetIdx, sheetName = sheetinfo
            sheet = excelFile.sheet_by_name(sheetName)
            itemNum = 1
            str1 = sheet.cell_value(3,2).strip()+sheet.cell_value(5,2).strip()
            str2 = sheet.cell_value(4,2).strip()

            dananhao = sheet.cell_value(1,0)
            dananhaoidx = dananhao.find(u"档案号：")
            dananhaolen = len(u"档案号：")
            if dananhaoidx == -1:
                dananhaoidx = dananhao.find(u"档案号:")
                dananhaolen = len(u"档案号:")
            dananhao = dananhao[dananhaoidx+dananhaolen:].strip()
            rowlen = 0
            for _ in sheet.get_rows():
                rowlen += 1
            for rowidx in xrange(7, rowlen+1):
                try:
                    if sheet.cell_value(rowidx,3).strip() == "":
                        break
                except:
                    pass
            zhangci = sheet.cell_value(rowidx - 1,3).strip()
            zhangciidx = zhangci.find("-")
            if zhangciidx == -1:
                zhangciidx = zhangci.find(u"－")
            zhangci = zhangci[zhangciidx+1:]

            modifyExcelFile.get_sheet(0).write(idx+3, 0, str(idx),stylenumber)
            modifyExcelFile.get_sheet(0).write(idx+3, 1, dananhao,styledanganhao)
            modifyExcelFile.get_sheet(0).write(idx+3, 2, str1+"\n"+str2,stylecontent)
            modifyExcelFile.get_sheet(0).write(idx+3, 6, zhangci,stylezhangci)
        except Exception as e:
            info.DisplayInfo("错误发生表单，错误表单名字如下：")
            info.DisplayInfo(sheetName)
            traceback.print_exc()
            raise e

    del excelFile

    try:
        modifyExcelFile.save(outputexcelFname)
    except:
        error.CanNotWrite(excelFname)
        common.ExitDueToError()

def autofillMain():
    # info.DisplayInfo( "请输入excel文件路径.")
    error.CreateLog("错误.txt")
    excelFName = sys.argv[1]
    outputexcelFName = sys.argv[2]
    info.DisplayInfo( "开始处理文件")
    ModifyExcel(excelFName,outputexcelFName)
    info.DisplayInfo( "处理文件结束")

if __name__ == "__main__":
    common.SafeRunProgram(autofillMain)
    # autofillMain()