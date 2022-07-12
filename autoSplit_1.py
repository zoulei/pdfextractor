# -*- coding:gbk -*-
#!/usr/bin/env python -u
# import extractRule
# import extractPDF
# import packfile
import os.path
import os
# import ConfigParser
import error
import common
# import modifyOrderNumber
import warnings
import sys
import traceback
import info
# import extractRule_1

import openpyxl
from pyPdf import PdfFileWriter, PdfFileReader
import shutil

def Split(all_page_config, pdf_dir, output_dir):
    pdf_dir = common.Path(pdf_dir)
    output_dir = common.Path(output_dir)
    if not os.path.exists(pdf_dir):
        info.DisplayInfo("目录不存在！路径：" + pdf_dir)
        return 1
    if not os.path.exists(output_dir):
        info.DisplayInfo("目录不存在！路径：" + output_dir)
        return 1
    for sheet_name, page_config in all_page_config.items():
        # print type(pdf_dir), type(sheet_name)
        pdfPath = os.path.join(pdf_dir, sheet_name+".pdf")
        if (not os.path.exists(pdfPath)):
            info.DisplayInfo("pdf文件不存在！路径：" + pdfPath)
            return 1
        inputPDF = PdfFileReader(open(pdfPath, "rb"))
        num_pages = inputPDF.getNumPages()
        # import pprint
        # pprint.pprint(page_config)
        if num_pages != page_config[-1][1] - 1:
            info.DisplayInfo(pdfPath + " 的页数为 " + str(num_pages) + ", excel中填的页数为 " + str(page_config[-1][1]))
            return 1
        sheet_dir = os.path.join(output_dir, sheet_name)

        if os.path.exists(sheet_dir):
            shutil.rmtree(sheet_dir)
        os.mkdir(sheet_dir)

        for i in range(len(page_config) - 1):
            write_fname = page_config[i][0].encode("gbk") + ".pdf"
            new_write_fname = common.CorrectFName(write_fname)
            if new_write_fname != write_fname:
                info.DisplayInfo("文件名有问题：" + write_fname)
            new_write_fname = os.path.join(sheet_dir, new_write_fname)
            # write_fname = os.path.join(sheet_dir, page_config[i][0].encode("gbk") + ".pdf")
            # print page_config[i][0].encode("gbk") + ".pdf"
            # new_write_fname = common.CorrectFName(write_fname)
            start_page = page_config[i][1]
            end_page = page_config[i + 1][1]
            output = PdfFileWriter()
            for j in range(start_page, end_page):
                output.addPage(inputPDF.getPage(j - 1))
            output.write(open(new_write_fname, "wb"))
    return 0

def ReadSplitConfig(excelFName):
    page_config = dict()
    excelFName = common.Path(excelFName)
    if not os.path.exists(excelFName):
        info.DisplayInfo("文件不存在！路径：" + excelFName)
        return 1
    excel_file = openpyxl.load_workbook(excelFName, True)
    all_sheet_names = excel_file.get_sheet_names()
    for sheet_name in all_sheet_names:
        sheet = excel_file.get_sheet_by_name(sheet_name)
        sheet_name = sheet_name.encode("gbk")
        page_config[sheet_name] = list()
        sheet_config = page_config[sheet_name]
        row = 3
        while True:
            out_fname = sheet[row][3].value
            if not out_fname:
                info.DisplayInfo(sheet_name + " 页处理失败，没有填文件题名，行号：" + str(row))
                return 1
            page = sheet[row][5].value
            if not page:
                info.DisplayInfo(sheet_name + " 页处理失败，页码填错，行号：" + str(row))
                return 1
            # print "page:", page
            page = str(page)
            page = "".join(page.split())
            page = page.replace(u"—", "-")
            if "-" in page:
                data = page.split("-")
                if len(data) != 2 or not data[0].isdigit() or not data[1].isdigit():
                    info.DisplayInfo(sheet_name + " 页处理失败，页码填错，行号：" + str(row))
                    return 1
                sheet_config.append([out_fname, int(data[0])])
                sheet_config.append(["", int(data[1]) + 1])
                break
            else:
                if not page.isdigit():
                    info.DisplayInfo(sheet_name + " 页处理失败，页码填错，行号：" + str(row))
                    return 1
                sheet_config.append([out_fname, int(page)])
            row += 1
        # dir(sheet)
        # print sheet_name, type(sheet_name), type(u"题名")
    return page_config

def AutoSplitMain():
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
    # error.CreateLog("错误.txt")

    excelFName = sys.argv[1]
    info.DisplayInfo("execel : " + excelFName)
    pdfDir = sys.argv[2]
    info.DisplayInfo("pdf dir : " + pdfDir)
    outputDir = sys.argv[3]
    info.DisplayInfo("output dir : " + outputDir)

    page_config = ReadSplitConfig(excelFName)
    if page_config != 1:
        if Split(page_config, pdfDir, outputDir) == 0:
            info.DisplayInfo("处理完成")
            return
    info.DisplayInfo("处理失败，请根据提示修正错误后重试")


if __name__ == "__main__":
    common.SafeRunProgram(AutoSplitMain)