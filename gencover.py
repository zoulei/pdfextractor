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

def ComputeMileLength(milestr):
    # print milestr
    mileData = milestr.split("+")
    start = mileData[0]
    end = mileData[1]
    for idx in xrange(len(start)):
        # print start[-idx],idx
        if not start[- idx - 1].isdigit():
            # print "notedi",start[-idx]
            break
        # print "isdi",start[-idx]
    # print end
    # print start[-idx+1:]
    diff = float(end) - float(start[-idx:])
    return diff

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
                # print type(res)
                # print len(res)
                # while True:
                #     print res[:1200]
                #     res = res[1200:]
                #     raw_input()
                end = res.rfind("</w:p>")
                start = res.find("<w:p")
                para = res[start:end+6]

                resultPara = ""
                # start = para.find(u"���ģ��һ")
                # print para[start:start + 500]

                for dataList in dataListList:
                    newPara = para
                    for idx,data in enumerate(dataList):
                        # print Constant.TEMPNAME+unicode(idx),data
                        newPara = newPara.replace(Constant.TEMPNAME+numberList[idx],data)
                        # print type(data)
                        # print type(newPara)
                    resultPara += newPara

                res = res[:start] + resultPara + res[end+6:]
                # for r in rep:
                #     res = res.replace(r,rep[r])
                buffer = res.encode("utf-8")
            zout.writestr(item, buffer)
        zout.close()
        zin.close()

def GenDocx(genDict):
    if not os.path.exists(Constant.COVERDIR):
        os.mkdir(Constant.COVERDIR)

    info.DisplayInfo("��ʼ��������")

    for fname in genDict:
        # time.sleep(10)

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
            zin = zipfile.ZipFile(Constant.DOCXTEMPFILE,"r")
            zout = zipfile.ZipFile(os.path.join(docdir,str(idx+3)+".docx"),"w")
            pass
        # print dataListList
            for item in zin.infolist():


                buffer = zin.read(item.filename)
                if (item.filename == 'word/document.xml'):
                    res = buffer.decode("utf-8")
                    # print type(res)
                    # print len(res)
                    # while True:
                    #     print res[:1200]
                    #     res = res[1200:]
                    #     raw_input()
                    # end = res.rfind("</w:p>")
                    # start = res.find("<w:p")
                    # para = res[start:end+6]

                    # resultPara = ""
                    # start = para.find(u"���ģ��һ")
                    # print para[start:start + 500]

                    # for dataList in dataListList:
                    # newPara = para
                    for idx,data in enumerate(dataList):
                        res = res.replace(Constant.TEMPNAME+numberList[idx],data)
                        # print Constant.TEMPNAME+unicode(idx),data
                        # newPara = newPara.replace(Constant.TEMPNAME+numberList[idx],data)
                        # print type(data)
                        # print type(newPara)
                    # resultPara += newPara

                    # res = res[:start] + resultPara + res[end+6:]
                    # for r in rep:
                    #     res = res.replace(r,rep[r])
                    buffer = res.encode("utf-8")
                zout.writestr(item, buffer)
            zout.close()
            zin.close()

def GenCoverXls(fname):
    info.DisplayInfo("��ʼ��ȡexcel�ļ��е���Ϣ")
    genDict = {}

    excelFile = xlrd.open_workbook(fname)
    sheetsName = excelFile.sheet_names()
    # print sheetsName
    # print type(sheetsName[0])
    # sheetsName[0].encode("gbk")
    # sheetsName = [str(v.encode("gbk")) for v in sheetsName]

    for sheetName in sheetsName:
        # error.ExcelMiddleItemError(sheetName.encode("gbk"),1 + 1)
        genDict[sheetName] = []
        sheet = excelFile.sheet_by_name(sheetName)
        for idx, row in enumerate(sheet.get_rows()):
            if idx <= 1:
                continue
            if sheet.cell(idx,0).value == "":
                continue

            try:
                # print sheet.cell(idx,2).value
                CData = [sheet.cell(idx,0).value.strip()]
                content = sheet.cell(idx,2).value.strip().split("\n")

                if len(content) < 3:
                    spacenum = 1

                    while( len(content) != 3):
                        contentstr = sheet.cell(idx,2).value.replace(u" "*spacenum,"\n")
                        contentstr = contentstr.replace(u"\u3000"*spacenum,"\n")
                        content = contentstr.strip().split("\n")
                        content = [v for v in content if v]
                        if len(content) <= 3:
                            break
                        spacenum += 1

                if len(content) != 3:
                    raise

                if content[1].find(u"��") == -1:
                    sep = u"("
                else:
                    sep = u"��"

                mile = content[1].split(sep)

                if content[1].find(u"��") == -1:
                    sep = u")"
                else:
                    sep = u"��"

                area = mile[1].split(sep)[0]
                # print ComputeMileLength(area),idx+1
                if ComputeMileLength(area) > 100:
                    double = True
                else:
                    double = False
                    # content[2] += u"���ϣ�"
                CData.extend([content[0],mile[0],area,content[2]])

                # return
                # CData.extend(sheet.cell(idx,2).value.strip().split("\n"))
                CData.append(sheet.cell(idx,3).value.strip())
                CData.append(sheet.cell(idx,6).value.strip())
                # print CData
                # return

                divide = sheet.cell(idx,5).value
                for ch in divide:
                    if ch == "2":
                        double == True
                        break
                    if ch == "1":
                        double = False
                        break

                if double:
                    raw = CData[4]
                    CData[4] += u"���ϣ�"
                    genDict[sheetName].append(CData)
                    CData = copy.deepcopy(CData)
                    CData[4] = raw + u"���£�"
                    genDict[sheetName].append(CData)
                else:
                    genDict[sheetName].append(CData)
                # print "note heaer"
                # print genDict[sheetName]
                # print CData
                # if idx > 300:
                #     break
            except:
                # print "===================="
                # print sheet.cell(idx,2).value.find(u" ")
                # print type(sheet.cell(idx,2).value)
                # print [sheet.cell(idx,2).value]
                # print sheet.cell(idx,2).value
                # print sheet.cell(idx,2).value.split("\n")
                # print "===================="
                error.ExcelMiddleItemError(sheetName.encode("gbk"),idx + 1)
                # print "��excel��%s�У���%d�з��������⣬��鿴." % (sheetName, idx + 1)
                # import traceback
                # traceback.print_exc()
                # print idx+1
                # if idx + 1 == 1257:
                #     break
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
    # excelFName = "F:/myfile/file/��ͼ/�Զ������/ʮ�߾��ƽ�Ŀ¼(7.27).xls"
    #excelFName = "F:\myfile/file\��ͼ\̷����/̷������� - ���� (2) - ����.xls"
    GenCover(excelFName)
    info.DisplayInfo("�������,����洢������·���У�\n"+ os.path.abspath(Constant.COVERDIR))
    error.Finish()

if __name__ == "__main__":
    common.SafeRunProgram(GenCoverMain)