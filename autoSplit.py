# -*- coding:gbk -*-
#!/usr/bin/env python -u
import extractRule
import extractPDF
import packfile
import os.path
import os
import ConfigParser
import error
import common
import modifyOrderNumber
import warnings
import sys
import traceback
import info
import extractRule_1


def autoSplit(excelFName,configData,pdfDir = None):
    excelFName = common.Path(excelFName)
    if pdfDir:
        dir = common.Path(pdfDir)
    else:
        dir = os.path.dirname(excelFName)
    info.DisplayInfo("开始计算提取方案")
    ruleDict = extractRule.ExtractExcel(excelFName,configData)
    # return
    info.DisplayInfo("开始提取PDF")
    for pdfName in ruleDict.keys():
        info.DisplayInfo("正在提取%s.pdf"%pdfName)
        try:
            pdfPath = os.path.join(dir,pdfName+".pdf")

            if (not os.path.exists(pdfPath)):
                error.PDFMiss(pdfPath)
                continue

            # the first 001 in /001/001
            outputDir = os.path.join(dir,pdfName)
            if not os.path.exists(outputDir):
                os.mkdir(outputDir)

            extractPDF.ExtractPDF(pdfPath,ruleDict[pdfName],outputDir)
            packfile.createFile(outputDir)
        except:
            error.PDFTranError(pdfName)
    # error.AddInfo("转换完毕，请根据提示更正excel中填错的项，如果excel文件中某个表单填写有错误，那么该页pdf的转换结果可能不准确，但是该错误不影响其他pdf的转换结果。")
    # error.Finish()
    # print '转换完毕，请查看  "错误.txt"  中的输出结果'
def checkSplitResult(excelFName, configData):
    excelFName = common.Path(excelFName)
    dir = os.path.dirname(excelFName)

    info.DisplayInfo("开始计算提取方案")
    ruleDict = extractRule_1.ExtractExcel(excelFName,configData)

    info.DisplayInfo("开始验证提取结果")
    for dirname in ruleDict.keys():
        info.DisplayInfo("正在检查%s.pdf"%dirname)

        targetdirname = os.path.join(dir,dirname)
        if (not os.path.exists(targetdirname)):
            error.DirMiss(targetdirname)
            continue
        extractPDF.CheckResult(ruleDict[dirname],targetdirname)

        # except:
        #     error.CheckSplitResultError(dirname)


# def InputConfig(excelFName):
#     if excelFName.endswith(".xls"):
#         excelFile = xlrd.open_workbook(excelFName)
#         sheetsName = excelFile.sheet_names()
#         sheetsName = [str(v) for v in sheetsName]
#     elif excelFName.endswith(".xlsx"):
#         excelFile = load_workbook(excelFName)
#         sheetsName = excelFile.get_sheet_names()
#         sheetsName = [str(v) for v in sheetsName]
#     else:
#         print "输入文件不是excel文件"
#         exit()
#         return None
#
#     coverDict = {}
#     ignoreDict = {}
#     repeatDict = {}
#
#     coverValue = "y"
def InputConfig(configFName):
    configFName = common.Path(configFName)
    coverDict = {}
    ignoreDict = {}
    repeatDict = {}

    parser = ConfigParser.SafeConfigParser(allow_no_value=True)
    parser.read(configFName)

    sectionList = parser.sections()
    for sectionName in sectionList:
        if sectionName == "cover":
            for key,value in parser.items(sectionName):
                if value:
                    try:
                        coverDict[key] = int(value)
                    except:
                        error.ConfigCover(key)

        if sectionName == "ignore":
            for key,value in parser.items(sectionName):
                if value:
                    try:
                        value = value.replace("，",",")
                        igdata = [int(v) for v in value.split(",")]
                        igdata.sort()
                        ignoreDict[key] = igdata
                    except:
                        error.ConfigIgnore(key)

        if sectionName == "repeat":
            for key,value in parser.items(sectionName):
                if value:
                    repeatDict[key] = {}
                    subDict = repeatDict[key]
                    try:
                        value = value.replace("，",",")
                        redata = [int(v) for v in value.split(",")]
                        redata.sort()
                        for page in redata:
                            if page not in subDict:
                                subDict[page] = -1
                            subDict[page] += 1
                    except:
                        error.ConfigRepeat(key)
    # info.DisplayInfo(coverDict)
    # info.DisplayInfo(repeatDict)
    return [coverDict,ignoreDict,repeatDict]

def AutoSplitMain():
    # print "请输入excel文件路径."
    # excelFName = raw_input()
    # print "请输入pdf文档所在的目录，如果pdf文档与excel上面输入的excel文件在同一个目录下，则可以跳过"
    # pdfDir = raw_input()
    # print "请输入配置文件路径，如果没有配置文件可以跳过"
    # configFName = raw_input()

    warnings.simplefilter("ignore")

    # excelFName = '"F:/myfile/file/鸿图/谭家湾/谭家湾大桥.xls"'
    # configFName = 'F:/myfile/file/鸿图/谭家湾/配置.txt'
    # pdfDir = ""
    # sys.stdout.write("sdgdsgdsg\n")
    info.DisplayInfo( "开始转换")
    # info.DisplayInfo( sys.argv)
    error.CreateLog("错误.txt")

    excelFName = sys.argv[1]
    pdfDir = sys.argv[2]
    if len(sys.argv) > 3:
        configFName = sys.argv[3]
    else:
        configFName = ""
    configData = []
    if configFName:
        info.DisplayInfo("正在读取配置信息")
        configData = InputConfig(configFName)
    info.DisplayInfo("正在调整excel的页号")
    modifyOrderNumber.ModifyExcel(excelFName)
    autoSplit(excelFName,configData,pdfDir)
    error.Finish()
    # info.DisplayInfo("提取完成")

if __name__ == "__main__":
    common.SafeRunProgram(AutoSplitMain)