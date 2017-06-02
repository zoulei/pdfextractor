# -*- coding:gbk -*-
import common
import info
import warnings
import error
import sys
import autoSplit
import modifyOrderNumber

def CheckSplitResultMain():
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
    # pdfDir = sys.argv[2]
    if len(sys.argv) > 2:
        configFName = sys.argv[2]
    else:
        configFName = ""
    configData = []
    if configFName:
        info.DisplayInfo("正在读取配置信息")
        configData = autoSplit.InputConfig(configFName)
    info.DisplayInfo("正在调整excel的页号")
    modifyOrderNumber.ModifyExcel(excelFName)
    autoSplit.checkSplitResult(excelFName,configData)
    error.FinishCheckResult()
    # info.DisplayInfo("提取完成")

if __name__ == "__main__":
    common.SafeRunProgram(CheckSplitResultMain)
