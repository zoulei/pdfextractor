# -*- coding:gbk -*-
import zipfile
import common
import sys
import info
import error
import xlrd
import os
import copy
import time
import Constant
import pathOperator

numberList = [u"һ",u"��",u"��",u"��",u"��",u"��",u"��",u"��",u"��",u"ʮ",u"ʮһ",
              u"ʮ��",u"ʮ��",u"ʮ��",u"ʮ��",u"ʮ��",u"ʮ��",u"ʮ��",u"ʮ��",u"��ʮ"]

def GenDocx1(genDict):
    if not os.path.exists(Constant.COVERDIR):
        os.mkdir(Constant.COVERDIR)

    info.DisplayInfo("��ʼ��������")

    for fname in genDict:
        # time.sleep(10)
        zin = zipfile.ZipFile(Constant.DOCXTEMPFILE,"r")
        info.DisplayInfo("��ʼ���������ļ��� "+fname.encode("gbk"))
        zout = zipfile.ZipFile(Constant.COVERDIR+"/"+fname.encode("gbk")+".docx","w")
        dataListList = genDict[fname]
        # print dataListList
        for item in zin.infolist():
            buffer = zin.read(item.filename)
            if (item.filename == 'word/document.xml'):
                res = buffer.decode("utf-8")
                end = res.rfind("</w:p>")
                start = res.find("<w:p")
                para = res[start:end+6]

                resultPara = ""

                for dataList in dataListList:
                    newPara = para
                    for idx,data in enumerate(dataList):
                        # print Constant.TEMPNAME+unicode(idx),data
                        newPara = newPara.replace(Constant.TEMPNAME+numberList[idx],data)
                    resultPara += newPara

                res = res[:start] + resultPara + res[end+6:]
                buffer = res.encode("utf-8")
            zout.writestr(item, buffer)
        zout.close()
        zin.close()

def GenDocx(genDict):
    if not os.path.exists(Constant.COVERDIR):
        os.mkdir(Constant.COVERDIR)

    info.DisplayInfo("��ʼ��������")

    for fname in genDict:
        info.DisplayInfo("��ʼ���������ļ��� "+fname.encode("gbk"))

        docdir = os.path.join(Constant.COVERDIR,fname.encode("gbk"))
        if not os.path.exists(docdir):
			os.mkdir(docdir)
        else:
            fileList = pathOperator.listfile(docdir)
            for delfilename in fileList:
                os.remove(delfilename)
        dataListList = genDict[fname]
        for idx,dataList in enumerate(dataListList):
            zin = zipfile.ZipFile(Constant.DOCXTEMPFILE1,"r")
            zout = zipfile.ZipFile(os.path.join(docdir,str(idx+3)+".docx"),"w")
            for item in zin.infolist():
                buffer = zin.read(item.filename)
                if (item.filename == 'word/document.xml'):
                    res = buffer.decode("utf-8")
                    for idx,data in enumerate(dataList):
                        res = res.replace(Constant.TEMPNAME+numberList[idx],data)
                    buffer = res.encode("utf-8")
                zout.writestr(item, buffer)
            zout.close()
            zin.close()

def GenCoverXls(fname):
    info.DisplayInfo("��ʼ��ȡexcel�ļ��е���Ϣ")
    genDict = {}

    excelFile = xlrd.open_workbook(fname)
    sheetsName = excelFile.sheet_names()
    lastidx = 0
    for sheetName in sheetsName:
        # error.ExcelMiddleItemError(sheetName.encode("gbk"),1 + 1)
        genDict[sheetName] = []
        sheet = excelFile.sheet_by_name(sheetName)
        for idx, row in enumerate(sheet.get_rows()):
            if idx <= 3:
                continue
            if sheet.cell(idx,0).value == "":
                if lastidx == 0:
                    break
                for inneridx in xrange(lastidx):
                    genDict[sheetName][-inneridx-1][-1] = str(lastidx)
                lastidx = 0
                continue
            try:
                # print sheet.cell(idx,2).value
                fileidx = int(sheet.cell(idx,0).value)
                content = sheet.cell(idx,2).value.strip().split("\n")
                firsttitle = content[0].replace("(",u"��")
                firsttitle = firsttitle.replace(")",u"��")
                firsttitle = firsttitle.replace(" ","")
                Kidx = firsttitle.find("K")
                leftkuoidx = firsttitle.find(u"��")
                rightkuoidx = firsttitle.find(u"��")

                title = firsttitle[:Kidx]
                workmile = firsttitle[leftkuoidx+1:rightkuoidx]
                completemile = firsttitle[Kidx:leftkuoidx]
                subtitle = content[1]

                company = sheet.cell(idx,11).value.strip()
                company = company.replace(u"��","")
                company = company.replace(u"��","")
                company = company.replace("(","")
                company = company.replace(")","")

                savetime = sheet.cell(idx,8).value.strip()
                code = sheet.cell(idx,1).value.strip()
                genDict[sheetName].append([company,title,completemile,workmile,subtitle,code,savetime,str(fileidx),str(fileidx)])
                lastidx = fileidx
            except:
                error.ExcelMiddleItemError(sheetName.encode("gbk"),idx + 1)
    info.DisplayInfo("excel�ļ��е���Ϣ��ȡ���")
    # print genDict
    GenDocx(genDict)
    # return genDict

def GenCover(fname):
    fname = common.Path(fname)
    if (fname.endswith(".xls")):
        return GenCoverXls(fname)
    elif (fname.endswith(".xlsx")):
        info.DisplayInfo("����֧��xlsx�ļ�������ϵ������Ա")
    else:
        error.NotExcelFile()
        # common.ExitDueToError()
        return None

def GenCoverMain():
    error.CreateLog("����.txt")
    excelFName = sys.argv[1]
    # excelFName = "F:/myfile/file/��ͼ/�½��ļ��� (2)/1 ����ƽ�Ŀ¼��������֣�1.xls"
    #excelFName = "F:\myfile/file\��ͼ\̷����/̷������� - ���� (2) - ����.xls"
    GenCover(excelFName)
    info.DisplayInfo("�������,����洢������·���У�\n"+ os.path.abspath(Constant.COVERDIR))
    error.Finish()

if __name__ == "__main__":
    common.SafeRunProgram(GenCoverMain)