# -*- coding:gbk -*-
# from openpyxl import load_workbook
import xlrd
import common
import error
import info

def ExtractFromTxt(fname):
    ruleList = []
    f = open(fname)
    idx = 0
    for line in f:
        pageCount = int(line.strip())
        ruleList.append(range(idx,idx + pageCount))
        idx += pageCount
    return ruleList

def ReadPage(pageStr):
    if isinstance(pageStr,unicode):
        pageStr = pageStr.encode("gbk")
    result = str(pageStr)
    dotIndex = result.find(".")
    if dotIndex != -1:
        result = result[:dotIndex]
    while(result.startswith("0")):
        result = result[1:]

    return result

# def ExtractXlsx(fname,configData):
#     splitDict = {}
#
#     excelFile = load_workbook(fname)
#     sheetsName = excelFile.get_sheet_names()
#     sheetsName = [str(v) for v in sheetsName]
#     for sheetName in sheetsName:
#         splitDict[sheetName] = {}
#         subDict = splitDict[sheetName]
#         sheet = excelFile.get_sheet_by_name(sheetName)
#
#         for idx, row in enumerate(sheet.rows):
#             if idx <= 3:
#                 continue
#             if row[0].value == None:
#                 continue
#             try:
#                 subDict[int(row[0].value)] = ReadPage(row[3].value)
#             except:
#                 error.ExcelMiddleItemError(sheetName,idx + 1)
#                 # print "在excel表单%s中，第%d行发生了问题，请查看." % (sheetName, idx + 1)
#
#     calSplitDict(splitDict, * configData)
#
#     return splitDict

def ExtractXls(fname,configData):
    splitDict = {}

    excelFile = xlrd.open_workbook(fname)
    sheetsName = excelFile.sheet_names()
    # sheetsName = [str(v) for v in sheetsName]
    for sheetName in sheetsName:
        # splitDict[sheetName] = {}
        # subDict = splitDict[sheetName]
        sheet = excelFile.sheet_by_name(sheetName)
        sheetName = sheetName.encode("gbk")

        try:
            rawdirname = sheet.cell(1,0).value
            if u":" in rawdirname:
                sep = u":"
            else:
                sep = u"："
            rawdirname = rawdirname.split(sep)[1]
        except:
            error.ExcelRowError(sheetName, 2)

        rawdirname = rawdirname.encode("gbk").strip()
        splitDict[rawdirname] = {}
        subDict = splitDict[rawdirname]

        for subconfigdict in configData:
            if sheetName in subconfigdict:
                subconfigdict[rawdirname] = subconfigdict[sheetName]

        for idx, row in enumerate(sheet.get_rows()):
            if idx <= 3:
                continue
            if sheet.cell(idx,0).value == "":
                continue
            if isinstance(sheet.cell(idx,0).value,unicode) and sheet.cell(idx,0).value.strip() == "":
                continue
            try:
                subDict[int(sheet.cell(idx,0).value)] = ReadPage(sheet.cell(idx,3).value)
            except:
                error.ExcelLineError(sheetName,idx + 1)
                # print "在excel表单%s中，第%d行发生了问题，请查看." % (sheetName, idx + 1)
                import traceback
                traceback.print_exc()

    calSplitDict(splitDict, * configData)

    return splitDict

def ExtractExcel(fname,configData):
    fname = common.Path(fname)
    if (fname.endswith(".xls")):
        return ExtractXls(fname,configData)
    elif (fname.endswith(".xlsx")):
        info.DisplayInfo("程序不支持xlsx文件，请联系开发人员")
        return None
        # return ExtractXlsx(fname,configData)
    else:
        error.NotExcelFile()
        common.ExitDueToError()
        return None

#splitDict :   sheetname : String => orderNumber : Int => List[Int] , list stores the page numbers of pdf to be extracted
#ignoreDict: sheetname: String => List[Int], list stores the page number that is ignored
#repeatDict: sheetname: String => repeatDict : Dict[pagenumber : Int => repeat numbers : Int]
#coverDict: sheetname: String => cover: boolean
def calSplitDict(splitDict, coverDict = None, ignoreDict = None, repeatDict = None):
    # import pprint
    # pp = pprint.PrettyPrinter(indent=4)
    # pp.pprint(repeatDict)

    # print "-"*100
    error.AddInfo("-"*100)

    keyList = splitDict.keys()
    # enter sheet
    for key in keyList:
        subDict = splitDict[key]
        subkeyList = subDict.keys()
        subkeyList.sort()
        if coverDict and coverDict.get(key) != None:
            cover = coverDict[key]
        else:
            cover = 0

        offset = 0
        if cover:
            offset += 1

        if ignoreDict and ignoreDict.get(key):
            ignoreList = ignoreDict[key]
        else:
            ignoreList = []
        ignoreList.sort()
        ignoreLength = len(ignoreList)
        ignoreIdx = 0

        if repeatDict and repeatDict.get(key):
            subRepeatDict = repeatDict[key]
        else:
            subRepeatDict = {}
        # pp.pprint(subRepeatDict)
        # print key
        repeatList = subRepeatDict.keys() # stores repeat page number
        repeatList.sort()
        repeatLength = len(repeatList)
        repeatIdx = 0

        pageMap = ConstructPageMap(subRepeatDict)
        # if key == "004":
        #     print subRepeatDict
        #     printList = pageMap.items()
        #     printList.sort(key = lambda x:x[1])
        #     for v in printList:
        #         print v
        #     exit()

        try:
            subkeyLength = len(subkeyList)
            # enter item of sheet
            for idx,subkey in enumerate(subkeyList):
                if idx != subkeyLength - 1:
                    startPage = subDict[subkey]
                else:
                    startPage, endPage = splitLastItem(subDict[subkeyList[-1]])
                    try:
                        # print endPage
                        endPage = NextPage(endPage, subRepeatDict)
                        # print endPage
                    except:
                        error.RepeatPagePostfix(key, subkey)

                if idx == subkeyLength - 2:
                    endPage = splitLastItem(subDict[subkeyList[-1]])[0]

                if idx < subkeyLength - 2:
                    endPage = subDict[subkey+1]

                currentIgnore = 0
                intStartPage = common.GetPageNumber(startPage)
                # print startPage,endPage
                intEndPage = common.GetPageNumber(endPage)
                while(ignoreIdx < ignoreLength):
                    ignorePage = ignoreList[ignoreIdx]
                    if ignorePage > intEndPage:
                        break
                    if ignorePage < intStartPage:
                        offset -= 1
                    if intStartPage < ignorePage < intEndPage:
                        currentIgnore -= 1
                    if ignorePage == intStartPage or ignorePage == intEndPage:
                        error.IgnorePageUsed(key,ignorePage,subkey)
                        # print "在excel表单%s中输入了第%d页漏打，但是该页码出现在第%d项或第%d项中.请检查excel文件." % (key,ignorePage,subkey, subkey+1)
                    ignoreIdx += 1

                # currentRepeat = 0
                # if key == "011" and subkey == 87:
                #     print repeatIdx
                # while(repeatIdx < repeatLength):
                #     repeatPage = repeatList[repeatIdx]
                #     if repeatPage > endPage:
                #         break
                #     if startPage < repeatPage == endPage:
                #         break
                #     if startPage <= repeatPage < endPage:
                #         currentRepeat += subRepeatDict[repeatPage]
                #     if startPage == repeatPage ==  endPage:
                #         currentRepeat += subRepeatDict[repeatPage]
                #     if repeatPage < startPage:
                #         print "在excel表单%s中输入了第%d页重复打码，但是该页没有被使用到.请检查excel文件" % (key,repeatPage)
                #     if repeatPage == startPage == endPage and currentRepeat > 1:
                #         print "在excel表单%s中包含了第%d页的项提取可能有误，请检查这些项的提取结果，如果有错误，请手动提取" % (key,repeatPage)
                #     repeatIdx += 1
                # if key == "011" and repeatIdx == 1:
                #     print "U"*100
                #     print subkey
                # if key == "011" and subkey == 1:
                #     print startPage,endPage
                #     print offset
                #     print currentIgnore
                #     print repeatList
                #     print repeatIdx
                startPage = GetMapedPage(startPage, pageMap)
                endPage = GetMapedPage(endPage, pageMap)
                # if key == "011" and subkey == 1:
                #     print startPage,endPage
                startPage = startPage + offset
                endPage = endPage + offset + currentIgnore
                offset += currentIgnore
                # offset += currentRepeat
                # if key == "011" and subkey == 1:
                #     print startPage,endPage
                subDict[subkey] = range(startPage  ,endPage )
        except:
            if idx == subkeyLength -1:
                error.ExcelLastItemError(key,subkey)
                # print "excel表单%s中，第%d项填得有问题，请查看." % (key, subkey)
            else:
                error.ExcelMiddleItemError(key,subkey)
                # print "excel表单%s中，第%d项或者第%d项填得有问题，请查看." % (key, subkey, subkey + 1)
            # import traceback
            # traceback.print_exc()
        # if key == "011":
        #     print subDict
            # import traceback
            # traceback.print_exc()
        # try:
        #     if len(subkeyList) == 0:
        #         continue
        #     lastStart,lastEnd = splitLastItem(subDict[subkeyList[-1]])
        #     subDict[subkeyList[-1]] = range(lastStart,lastEnd+1)
        # except:
        #     print "excel表单%s中，倒数第1项或者倒数第2项填得有问题，请查看." % key

        for idx, subkey in enumerate(subkeyList):
            if not subDict[subkey]:
                error.EmptyItem(key,subkey)
                # print "excel表单%s中，第%d项或者其周围的项填得有问题，或者存在重复打码情况没有在配置文件中没有注明，请查看." % (key, subkey)

    # print "="*100
    error.AddInfo("="*100)
    # print splitDict["011"]
    # exit()

def ConstructPageMap(subRepeatDict):
    offset = 0
    repeatList = subRepeatDict.keys()
    repeatList.sort()
    lastRepeatPage = 0

    pageMap = {}
    # if len(repeatList) == 2:
    #     print repeatList
    for repeatPage in repeatList:
        for v in xrange(lastRepeatPage + 1,repeatPage):
            # if len(repeatList) == 2:
            #     print v,v + offset
            pageMap[str(v)] = v + offset
        repeatCount = subRepeatDict[repeatPage]
        for idx in xrange(repeatCount + 1):
            pageMap[str(repeatPage) + chr(ord("A") + idx)] = repeatPage + idx + offset
        offset += repeatCount
        lastRepeatPage = repeatPage

    pageMap["offset"] = offset
    return pageMap

def NextPage(curPage, repeatDict):
    if (curPage[-1].isdigit() ) :
        curPage = int(curPage)
        if curPage in repeatDict:
            raise
        curPage += 1
        if curPage in repeatDict:
            return str(curPage) + "A"
        else:
            return str(curPage)
    else:
        if (common.SubstractBA(curPage[-1],"A") < (repeatDict[int(curPage[:-1])] -1) ):
            return curPage[:-1] + common.Aadd1(curPage[-1],1)
        else:
            if (int(curPage[:-1]) + 1) in repeatDict:
                return curPage[:-1] + "A"
            else:
                return str(int(curPage[:-1]) + 1)

def GetMapedPage(pageStr, pageMap):
    if pageStr in pageMap:
        return pageMap[pageStr]
    else:
        return pageMap["offset"] + int(pageStr)

def splitLastItem(str):
    if str.find("-") != -1:
        sep = "-"
    elif str.find("—") != -1:
        sep = "—"
    elif str.find("~") != -1:
        sep = "~"
    elif str.find("～") != -1:
        sep = "～"
    lastStart,lastEnd = str.split(sep)
    return [lastStart,lastEnd]

if __name__ == "__main__":
    error.CreateLog("错误.txt")
    excelName = raw_input()
    ExtractExcel(excelName,[])