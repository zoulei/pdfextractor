# -*- coding: gbk -*-

import os.path
import os
import traceback
import sys

def Path(fname):
    if fname.endswith("\""):
        fname = fname[1:-1]
    return fname

def Aadd1(char, number):
    return chr(ord(char) + number)

def SubstractBA(charB,charA):
    return ord(charB) - ord(charA)

def GetPageNumber(pageStr):
    if pageStr[-1].isdigit():
        return int(pageStr)
    else:
        return int(pageStr[:-1])

def GeneraterPageNumber(pageStr):
    dotIndex = pageStr.find(".")
    if dotIndex != -1:
        pageStr = pageStr[:dotIndex]

    iterIdx = 0
    while(iterIdx < len(pageStr)):
        if pageStr[iterIdx].isalpha():
            pageStr = pageStr[:iterIdx] + pageStr[iterIdx + 1:]
        else:
            iterIdx += 1
    # for idx in xrange(len(pageStr)):
    #     if pageStr[- idx - 1].isalpha():
    #         pageStr = pageStr[: (- idx -1)] + pageStr[(- idx):]
    while len(pageStr) < 3:
        pageStr = "0" + pageStr
    return pageStr

def RenameFile(sourceName, targetName):
    if os.path.exists(targetName):
        os.remove(targetName)
    os.rename(sourceName, targetName)

def ExitDueToError():
    traceback.print_exc()
    # raw_input()
    print "程序发生错误，请保存错误提示并联系开发人员"
    try:
        sys.exit()
    except SystemExit:
        os._exit(1)


def SafeRunProgram(func, args = []):
    try:
        func(*args)
    except:
        ExitDueToError()

def ProDir(func,dirName):
    fileList = []
    tempDir = [dirName]
    while(tempDir):
        dirName = tempDir.pop()
        fileNameList = os.listdir(dirName)
        for fname in fileNameList:
            pathName = os.path.join(dirName,fname)
            if os.path.isfile(pathName):
                fileList.append(pathName)
            else:
                tempDir.append(pathName)
    for fname in fileList:
        func(fname)

if __name__ == "__main__":
    print GeneraterPageNumber("081-082")
    print GeneraterPageNumber("081-082A")
    print GeneraterPageNumber("081A-082")
    print GeneraterPageNumber("1.0")
    print GeneraterPageNumber("001")