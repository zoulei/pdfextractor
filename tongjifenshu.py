# -*- coding:gbk -*-

from openpyxl import load_workbook
from xlutils.filter import process,XLRDReader, XLWTWriter
import xlrd
import xlutils.copy
import configuration
import error
import common
import Constant
import info
import os
import math
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

def ModifyXlsx(excelFName):
    excelFile = load_workbook(excelFName,read_only=False)
    sheetsName = excelFile.get_sheet_names()
    sheetsName = [str(v) for v in sheetsName]

    for sheetName in sheetsName:
        itemNum = 1
        sheet = excelFile.get_sheet_by_name(sheetName)
        for idx, row in enumerate(sheet.rows):
            if idx < configuration.STARTROW - 1:
                continue
            if row[3].value == None:
                continue
            pos = "%s%d"%(configuration.GetChar(configuration.ORDERCOL),idx + 1)
            sheet[pos] = itemNum
            sheet[pos].style = row[3].style
            itemNum += 1

    try:
        excelFile.save(excelFName)
    except:
        error.CanNotWrite(excelFName)
        common.ExitDueToError()

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

    # formatFile = xlrd.open_workbook(Constant.FROMATFILE,formatting_info=True)
    # formatMo,formatStype = copy2(formatFile)
    #
    # # formatFile.sheet_by_index(0)
    # xf_index = formatFile.sheet_by_index(0).cell_xf_index(6,0)
    # saved_style = formatStype[xf_index]
    #
    # xf_index = formatFile.sheet_by_index(0).cell_xf_index(6,3)
    # pageStyle = formatStype[xf_index]

    score304Map = {}

    score417Map = {}

    sheetsName = excelFile.sheet_names()
    # sheetsName = [str(v.encode("gbk")) for v in sheetsName]
    for sheetIdx,sheetName in enumerate(sheetsName):
        sheet = excelFile.sheet_by_name(sheetName)
        itemNum = 1

        # print sheetName
        for idx, row in enumerate(sheet.get_rows()):
            # if sheetName == "479":
            #     print idx,row[3].value, "--", itemNum
            if idx == 0:
                continue

            if idx >30:
                break

            colidx = 12
            print len(row),sheetName,idx, excelFname
            if len(row) > colidx + 1:
                if row[colidx + 1].value != "":
                    colidx += 1

            if len(row) <= colidx:
                continue

            if row[colidx].value == "":
                continue

            if sheetIdx == 0:
                # 304
                score304Map[idx] = float(row[colidx].value)
            else:
                score417Map[idx] = float(row[colidx].value)

            # if saved_style == None:
            #     print "skdgsdgsdgsdgdsgsdgdsg"*100
            # modifyExcelFile.get_sheet(sheetIdx).write(idx,configuration.ORDERCOL - 1,itemNum,saved_style)
            # modifyExcelFile.get_sheet(sheetIdx).write(idx,configuration.PAGECOL - 1,row[3].value,pageStyle)
            itemNum += 1

    return [score304Map,score417Map]
    # del excelFile
    #
    # try:
    #     modifyExcelFile.save(excelFname)
    # except:
    #     error.CanNotWrite(excelFname)
    #     common.ExitDueToError()

def WriteScoreXls(excelFname,scorelist):
    excelFile = xlrd.open_workbook(excelFname,formatting_info=True)
    modifyExcelFile,outStyle = copy2(excelFile)

    offsetlist = []
    for scoremap in scorelist:
        offsetlist.append(math.ceil(sum(scoremap.values()) / len(scoremap) - 89) )
    sheetsName = excelFile.sheet_names()
    # sheetsName = [str(v.encode("gbk")) for v in sheetsName]
    print offsetlist
    for idx,singleclassscore in enumerate(scorelist):
        for key ,value in singleclassscore.items():
            modifyExcelFile.get_sheet(idx).write(key,12,value-offsetlist[idx])

    del excelFile

    try:
        modifyExcelFile.save(excelFname)
    except:
        error.CanNotWrite(excelFname)
        common.ExitDueToError()

def tongjiscore(dirname):
    dirname = common.Path(dirname)
    flist = os.listdir(dirname)

    scorelist = []

    for fname in flist:
        if fname == "总成绩.xls":
            continue
        fname = os.path.join(dirname,fname)
        scorelist.append(ModifyExcel(fname) )

    totalscore = [{},{}]
    for singlescore in scorelist:
        for idx,singlescoremap in enumerate(singlescore):
            for key,value in singlescoremap.items():
                if key not in totalscore[idx]:
                    totalscore[idx][key] = 0
                totalscore[idx][key] += value

    WriteScoreXls(os.path.join(dirname,"总成绩.xls"),totalscore)

if __name__ == "__main__":
    dirfname = raw_input()
    tongjiscore(dirfname)
