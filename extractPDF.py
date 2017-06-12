# -*- coding:gbk -*-

import extractRule
from pyPdf import PdfFileWriter, PdfFileReader
import os.path
import traceback
import error

def CheckResult(rule,outputDir):
    keyList = rule.keys()

    for idx in keyList:
        pageIdxList = rule[idx]
        count = len(pageIdxList)

        outputFName = str(idx)
        if (len(outputFName) < 3):
            outputFName = "0"*(3-len(outputFName)) + outputFName
        outputFName = os.path.join(outputDir,outputFName,outputFName+".pdf")

        if not os.path.exists(outputFName):
            error.PDFMiss(outputFName)
            continue

        inputPDF = PdfFileReader(open(outputFName,"rb"))
        if count != inputPDF.numPages:
            error.PageNotMatch(outputFName)

def ExtractPDF(sourcePDF, rule, outputDir):
    inputPDF = PdfFileReader(open(sourcePDF,"rb"))

    keyList = rule.keys()

    for idx in keyList:
        pageIdxList = rule[idx]
        output = PdfFileWriter()
        for pageIdx in pageIdxList:
            try:
                output.addPage(inputPDF.getPage(pageIdx-1))
            except:
                error.PDFPageMiss(sourcePDF,pageIdx,idx)
                # print "pdf文件:"+sourcePDF+"中第" + str(pageIdx) + "页不存在或者无法读取.请查看pdf文件和excel文件中第"+str(idx)+"项的页码编号.或者文档存在漏打码的情况。"
                # print pageIdxList
                # traceback.print_exc()

        outputFName = str(idx)
        if (len(outputFName) < 3):
            outputFName = "0"*(3-len(outputFName)) + outputFName
        outputFName = os.path.join(outputDir,outputFName+".pdf")
        output.write(open(outputFName,"wb"))
        # output.close()

    keyList.sort()
    lastPage = rule[keyList[-1]]
    lastPage.sort()
    if (lastPage[-1] < inputPDF.getNumPages()):
        error.PDFPageNotUsed(sourcePDF)
        # print "pdf文件：%s中的最后几页没有被使用到，可能是excel表中项填写不完整，也可能是文档存在重复打码的情况。"%sourcePDF

    # inputPDF.close()

def DeletePage():
    inputPDF = PdfFileReader(open("E:/迅雷下载/聂怡静毕业论文20170520.pdf","rb"))
    output = PdfFileWriter()
    for idx in xrange(43):
        if idx == 3:
            continue
        output.addPage(inputPDF.getPage(idx))
    output.write(open("tmp.pdf","wb"))

if __name__ == "__main__":
    # raw_input("why")
    # ExtractPDF("C:/Users/zoulei/PycharmProjects/pdfextractor/20151028-Opprentice (IMC)-2.pdf",extractRule.ExtractFromTxt("C:/Users/zoulei/PycharmProjects/pdfextractor/123.txt"),"C:/Users/zoulei/PycharmProjects/pdfextractor/output")
    # raw_input("press any key to exit")
    DeletePage()