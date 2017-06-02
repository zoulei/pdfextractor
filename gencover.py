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

numberList = [u"一",u"二",u"三",u"四",u"五",u"六",u"七",u"八",u"九",u"十",u"十一",
              u"十二",u"十三",u"十四",u"十五",u"十六",u"十七",u"十八",u"十九",u"二十"]

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

    info.DisplayInfo("开始制做封面")

    for fname in genDict:
        # time.sleep(10)
        zin = zipfile.ZipFile(Constant.DOCXTEMPFILE,"r")
        info.DisplayInfo("开始制做封面文件： "+fname.encode("gbk"))
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
                # start = para.find(u"填充模板一")
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

    info.DisplayInfo("开始制做封面")

    for fname in genDict:
        # time.sleep(10)

        info.DisplayInfo("开始制做封面文件： "+fname.encode("gbk"))

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
                    # start = para.find(u"填充模板一")
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
    info.DisplayInfo("开始提取excel文件中的信息")
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

                if content[1].find(u"（") == -1:
                    sep = u"("
                else:
                    sep = u"（"

                mile = content[1].split(sep)

                if content[1].find(u"）") == -1:
                    sep = u")"
                else:
                    sep = u"）"

                area = mile[1].split(sep)[0]
                # print ComputeMileLength(area),idx+1
                if ComputeMileLength(area) > 100:
                    double = True
                else:
                    double = False
                    # content[2] += u"（上）"
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
                    CData[4] += u"（上）"
                    genDict[sheetName].append(CData)
                    CData = copy.deepcopy(CData)
                    CData[4] = raw + u"（下）"
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
                # print "在excel表单%s中，第%d行发生了问题，请查看." % (sheetName, idx + 1)
                # import traceback
                # traceback.print_exc()
                # print idx+1
                # if idx + 1 == 1257:
                #     break
    info.DisplayInfo("excel文件中的信息提取完毕")
    # print genDict
    GenDocx(genDict)
    # return genDict

def GenCover(fname):
    fname = common.Path(fname)
    if (fname.endswith(".xls")):
        return GenCoverXls(fname)
    elif (fname.endswith(".xlsx")):
        info.DisplayInfo("程序不支持xlsx文件，请联系开发人员")
    else:
        error.NotExcelFile()
        # common.ExitDueToError()
        return None

def GenCoverMain():
    error.CreateLog("错误.txt")
    excelFName = sys.argv[1]
    # excelFName = "F:/myfile/file/鸿图/自动打封面/十七局移交目录(7.27).xls"
    #excelFName = "F:\myfile/file\鸿图\谭家湾/谭家湾大桥 - 副本 (2) - 副本.xls"
    GenCover(excelFName)
    info.DisplayInfo("处理完毕,结果存储在如下路径中：\n"+ os.path.abspath(Constant.COVERDIR))
    error.Finish()

if __name__ == "__main__":
    common.SafeRunProgram(GenCoverMain)