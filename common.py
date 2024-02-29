# -*- coding: gbk -*-

import os.path
import os
import traceback
import sys
import re

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
    print traceback.format_exc()
    # raw_input()
    print "程序发生错误，请联系开发人员"
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

def CorrectFName(fname):
    for v in "\\/:*?\"<>|":
        fname = fname.replace(v, "")
    return fname

def FindFNameByIdx(fname_list, idx):
    for fname in fname_list:
        first_idx = fname.find("-")
        if first_idx == -1:
            continue
        last_pos = len(fname)
        for i in range(first_idx + 1, len(fname)):
            if not fname[i].isdigit():
                last_pos =  i
                break
        idx_str = fname[first_idx + 1: last_pos]
        if not idx_str:
            continue
        if int(idx_str) == idx:
            return fname
    return ""

if __name__ == "__main__":
    print GeneraterPageNumber("081-082")
    print GeneraterPageNumber("081-082A")
    print GeneraterPageNumber("081A-082")
    print GeneraterPageNumber("1.0")
    print GeneraterPageNumber("001")