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
    info.DisplayInfo("ת����ϡ�")

def Finish():
    global errorLog

    if errorLog.HasError():
        AddInfo("������ϣ�����һЩ���⣬�������ʾ����excel������������ڸ���������������г���")
        info.DisplayInfo('ת����ϣ���鿴  "����.txt"  �е�������')
    else:
        AddInfo("������ϣ�û�з��ִ��󡣵��Ⲣ����֤���еĽ������ȷ��")
        info.DisplayInfo( "ת����ϣ�û�з��ִ��󡣵��Ⲣ����֤���еĽ������ȷ��")
    errorLog.close()

def FinishCheckResult():
    global errorLog

    if errorLog.HasError():
        AddInfo("�����ϣ�����һЩ���⣬�������ʾ����excel������������ڸ���������������г���")
        info.DisplayInfo('�����ϣ���鿴  "����.txt"  �е�������')
    else:
        AddInfo("�����ϣ�����û�з��ִ���")
        info.DisplayInfo( "�����ϣ�����û�з��ִ���")
    errorLog.close()

def ConfigCover(str):
    log = "�����ļ����Ƿ��з����������%s���ʽ�д�����" % str
    AddErrorLog(log)

def ConfigIgnore(str):
    log = "�����ļ���©�����������%s���ʽ�д�����" % str
    AddErrorLog(log)

def ConfigRepeat(str):
    log = "�����ļ����ظ������������%s���ʽ�д�����" % str
    AddErrorLog(log)

def NotExcelFile():
    log = "�����ļ�����excel�ļ�"
    AddErrorLog(log)

def RepeatPagePostfix(sheet,item):
    log = "excel��%s�е�%d����д�����ظ�����ҳ�ź���û�м���ABCD" % (sheet, item)
    AddErrorLog(log)

def PDFMiss(pdfName):
    log ="pdf�ļ�:" + pdfName + "    ������"
    AddErrorLog(log)

def DirMiss(dirname):
    log ="Ŀ¼:" + dirname + "    ������"
    AddErrorLog(log)

def PDFTranError(pdfName):
    log = "%sת���д�" % pdfName
    AddErrorLog(log)

def PDFPageMiss(sourcePDF, pageIdx, item):
    log = "pdf�ļ�:"+sourcePDF+"�е�" + str(pageIdx) + "ҳ�����ڻ����޷���ȡ.��鿴pdf�ļ���excel�ļ��е�"+str(item)+"���ҳ����.�����ĵ�����©����������"
    AddErrorLog(log)

def PDFPageNotUsed(sourcePDF):
    log = "pdf�ļ���%s�е����ҳû�б�ʹ�õ���������excel��������д��������Ҳ�������ĵ������ظ�����������"%sourcePDF
    AddErrorLog(log)

def ExcelRowError(sheetName, item):
    log = "��excel��%s�У���%d�з��������⣬��鿴." % (sheetName, item)
    AddErrorLog(log)

def ExcelLastItemError(sheetName, item):
    log = "��excel��%s�У���%d�з��������⣬��鿴." % (sheetName, item)
    AddErrorLog(log)

def ExcelMiddleItemError(sheetName, item):
    log = "excel��%s�У���%d����ߵ�%d����������⣬��鿴." % (sheetName, item, item + 1)
    AddErrorLog(log)

def ExcelLineError(sheetName, item):
    log = "excel��%s�У���%d�л��ߵ�%d����������⣬��鿴." % (sheetName, item, item + 1)
    AddErrorLog(log)

def IgnorePageUsed(sheet, ignorePage, item):
    log = "��excel��%s�������˵�%dҳ©�򣬵��Ǹ�ҳ������ڵ�%d����%d����.����excel�ļ�." % (sheet,ignorePage,item, item+1)
    AddErrorLog(log)

def EmptyItem(sheet,item):
    log = "excel��%s�У���%d���������Χ������������⣬���ߴ����ظ��������û���������ļ���ע������鿴." % (sheet, item)
    AddErrorLog(log)

def CheckSplitResultError(dirname):
    log = "��� %s ����ȡ���ʱ����������" % dirname
    AddErrorLog(log)

def PageNotMatch(fname):
    log = "�ļ���%s ��ƥ��" % fname
    AddErrorLog(log)

def RotateError(fname, pagenum):
    log = "�ļ� %s �ĵ� %d ҳ����ʶ���������˹�����ҳ��ʶ����������" % (fname, pagenum)
    AddErrorLog(log)

def AddInfo(info):
    log = info
    AddInfoLog(log)

#==========================================================================================
def CanNotWrite(fname):
    msg = '�޷����ļ�   "%s"   ��д����Ϣ�������Ǹ��ļ��Ѿ�����������򿪡�'%fname
    info.DisplayInfo( msg)
