# -*- coding:gbk -*-
import common
import info
import warnings
import error
import sys
import autoSplit
import modifyOrderNumber

def CheckSplitResultMain():
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
    # pdfDir = sys.argv[2]
    if len(sys.argv) > 2:
        configFName = sys.argv[2]
    else:
        configFName = ""
    configData = []
    if configFName:
        info.DisplayInfo("���ڶ�ȡ������Ϣ")
        configData = autoSplit.InputConfig(configFName)
    info.DisplayInfo("���ڵ���excel��ҳ��")
    modifyOrderNumber.ModifyExcel(excelFName)
    autoSplit.checkSplitResult(excelFName,configData)
    error.FinishCheckResult()
    # info.DisplayInfo("��ȡ���")

if __name__ == "__main__":
    common.SafeRunProgram(CheckSplitResultMain)
