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
    info.DisplayInfo("��ʼ������ȡ����")
    ruleDict = extractRule.ExtractExcel(excelFName,configData)
    # return
    info.DisplayInfo("��ʼ��ȡPDF")
    for pdfName in ruleDict.keys():
        info.DisplayInfo("������ȡ%s.pdf"%pdfName)
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
    # error.AddInfo("ת����ϣ��������ʾ����excel����������excel�ļ���ĳ������д�д�����ô��ҳpdf��ת��������ܲ�׼ȷ�����Ǹô���Ӱ������pdf��ת�������")
    # error.Finish()
    # print 'ת����ϣ���鿴  "����.txt"  �е�������'
def checkSplitResult(excelFName, configData):
    excelFName = common.Path(excelFName)
    dir = os.path.dirname(excelFName)

    info.DisplayInfo("��ʼ������ȡ����")
    ruleDict = extractRule_1.ExtractExcel(excelFName,configData)

    info.DisplayInfo("��ʼ��֤��ȡ���")
    for dirname in ruleDict.keys():
        info.DisplayInfo("���ڼ��%s.pdf"%dirname)

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
#         print "�����ļ�����excel�ļ�"
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
                        value = value.replace("��",",")
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
                        value = value.replace("��",",")
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
    # print "������excel�ļ�·��."
    # excelFName = raw_input()
    # print "������pdf�ĵ����ڵ�Ŀ¼�����pdf�ĵ���excel���������excel�ļ���ͬһ��Ŀ¼�£����������"
    # pdfDir = raw_input()
    # print "�����������ļ�·�������û�������ļ���������"
    # configFName = raw_input()

    warnings.simplefilter("ignore")

    # excelFName = '"F:/myfile/file/��ͼ/̷����/̷�������.xls"'
    # configFName = 'F:/myfile/file/��ͼ/̷����/����.txt'
    # pdfDir = ""
    # sys.stdout.write("sdgdsgdsg\n")
    info.DisplayInfo( "��ʼת��")
    # info.DisplayInfo( sys.argv)
    error.CreateLog("����.txt")

    excelFName = sys.argv[1]
    pdfDir = sys.argv[2]
    if len(sys.argv) > 3:
        configFName = sys.argv[3]
    else:
        configFName = ""
    configData = []
    if configFName:
        info.DisplayInfo("���ڶ�ȡ������Ϣ")
        configData = InputConfig(configFName)
    info.DisplayInfo("���ڵ���excel��ҳ��")
    modifyOrderNumber.ModifyExcel(excelFName)
    autoSplit(excelFName,configData,pdfDir)
    error.Finish()
    # info.DisplayInfo("��ȡ���")

if __name__ == "__main__":
    common.SafeRunProgram(AutoSplitMain)