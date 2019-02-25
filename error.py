# -*- coding:gbk -*-
import info
errorLog = None

class ErrorLog:
    def __init__(self, fname):
        self.m_fname = fname
        self.m_file = None
        self.m_Err = 0
        self.m_Info = 0
        self.m_Len = 0

        self.m_file = open(self.m_fname, "w")

    def add(self, log):
        if not self.m_file:
            self.m_file = open(self.m_fname, "w")
        self.m_file.write(log + "\n")
        self.m_file.flush()

    def close(self):
        if self.m_file:
            self.m_file.close()
            self.m_file = None

    def addError(self,errMsg):
        self.add(errMsg)
        self.m_Err += 1
        self.m_Len += 1

    def addInfo(self,infoMsg):
        self.add(infoMsg)
        self.m_Info += 1
        self.m_Len += 1

    def HasError(self):
        return bool(self.m_Err)

    def HasInfo(self):
        return bool(self.m_Info)

    def __del__(self):
        self.close()

def CreateLog(fname):
    global errorLog

    if not errorLog:
        errorLog = ErrorLog(fname)

def AddErrorLog(log):
    global errorLog

    errorLog.addError(log)

def AddInfoLog(log):
    global errorLog

    errorLog.addInfo(log)

def RealFinish():
    global errorLog
    info.DisplayInfo("转换完毕。")

def Finish():
    global errorLog

    if errorLog.HasError():
        AddInfo("处理完毕，发现一些问题，请根据提示更正excel中填错的项，并且在更正错误后重新运行程序。")
        info.DisplayInfo('转换完毕，请查看  "错误.txt"  中的输出结果')
    else:
        AddInfo("处理完毕，没有发现错误。但这并不保证所有的结果都正确。")
        info.DisplayInfo( "转换完毕，没有发现错误。但这并不保证所有的结果都正确。")
    errorLog.close()

def FinishCheckResult():
    global errorLog

    if errorLog.HasError():
        AddInfo("检查完毕，发现一些问题，请根据提示更正excel中填错的项，并且在更正错误后重新运行程序。")
        info.DisplayInfo('检查完毕，请查看  "错误.txt"  中的输出结果')
    else:
        AddInfo("检查完毕，程序没有发现错误。")
        info.DisplayInfo( "检查完毕，程序没有发现错误。")
    errorLog.close()

def ConfigCover(str):
    log = "配置文件中是否含有封面配置项第%s项格式有错，请检查" % str
    AddErrorLog(log)

def ConfigIgnore(str):
    log = "配置文件中漏打码配置项第%s项格式有错，请检查" % str
    AddErrorLog(log)

def ConfigRepeat(str):
    log = "配置文件中重复打码配置项第%s项格式有错，请检查" % str
    AddErrorLog(log)

def NotExcelFile():
    log = "输入文件不是excel文件"
    AddErrorLog(log)

def RepeatPagePostfix(sheet,item):
    log = "excel表单%s中第%d项填写有误，重复打码页号后面没有加上ABCD" % (sheet, item)
    AddErrorLog(log)

def PDFMiss(pdfName):
    log ="pdf文件:" + pdfName + "    不存在"
    AddErrorLog(log)

def DirMiss(dirname):
    log ="目录:" + dirname + "    不存在"
    AddErrorLog(log)

def PDFTranError(pdfName):
    log = "%s转换有错" % pdfName
    AddErrorLog(log)

def PDFPageMiss(sourcePDF, pageIdx, item):
    log = "pdf文件:"+sourcePDF+"中第" + str(pageIdx) + "页不存在或者无法读取.请查看pdf文件和excel文件中第"+str(item)+"项的页码编号.或者文档存在漏打码的情况。"
    AddErrorLog(log)

def PDFPageNotUsed(sourcePDF):
    log = "pdf文件：%s中的最后几页没有被使用到，可能是excel表中项填写不完整，也可能是文档存在重复打码的情况。"%sourcePDF
    AddErrorLog(log)

def ExcelRowError(sheetName, item):
    log = "在excel表单%s中，第%d行发生了问题，请查看." % (sheetName, item)
    AddErrorLog(log)

def ExcelLastItemError(sheetName, item):
    log = "在excel表单%s中，第%d行发生了问题，请查看." % (sheetName, item)
    AddErrorLog(log)

def ExcelMiddleItemError(sheetName, item):
    log = "excel表单%s中，第%d项或者第%d项填得有问题，请查看." % (sheetName, item, item + 1)
    AddErrorLog(log)

def ExcelLineError(sheetName, item):
    log = "excel表单%s中，第%d行或者第%d行填得有问题，请查看." % (sheetName, item, item + 1)
    AddErrorLog(log)

def IgnorePageUsed(sheet, ignorePage, item):
    log = "在excel表单%s中输入了第%d页漏打，但是该页码出现在第%d项或第%d项中.请检查excel文件." % (sheet,ignorePage,item, item+1)
    AddErrorLog(log)

def EmptyItem(sheet,item):
    log = "excel表单%s中，第%d项或者其周围的项填得有问题，或者存在重复打码情况没有在配置文件中注明，请查看." % (sheet, item)
    AddErrorLog(log)

def CheckSplitResultError(dirname):
    log = "检查 %s 的提取结果时程序发生错误" % dirname
    AddErrorLog(log)

def PageNotMatch(fname):
    log = "文件：%s 不匹配" % fname
    AddErrorLog(log)

def RotateError(fname, pagenum):
    log = "文件 %s 的第 %d 页可能识别有误，请人工检查该页的识别结果并调整" % (fname, pagenum)
    AddErrorLog(log)

def AddInfo(info):
    log = info
    AddInfoLog(log)

#==========================================================================================
def CanNotWrite(fname):
    msg = '无法往文件   "%s"   中写入信息，可能是该文件已经被其他程序打开。'%fname
    info.DisplayInfo( msg)
