# -*- coding:gbk -*-
#!/usr/bin/env python -u

import os.path
import os
import common
import warnings
import sys
import info

import openpyxl
from PyPDF2 import PdfFileWriter, PdfFileReader
from win32com import client

def Split(excelFName, beikaobiao_fname, pdf_dir, output_dir, appName):
    pdf_dir = common.Path(pdf_dir)
    output_dir = common.Path(output_dir)
    excelFName = common.Path(excelFName)
    beikaobiao_fname = common.Path(beikaobiao_fname)
    excel = client.Dispatch(appName)
    # excel = client.Dispatch("Excel.Application")
    # excel = client.gencache.EnsureDispatch("et.Application")
    # print "excel : ", excel
    sheets = excel.Workbooks.Open(excelFName)

    beikaobiao_wb = excel.Workbooks.Open(beikaobiao_fname)
    beikaobiao_sheet = beikaobiao_wb.Worksheets[0]
    for tmp_sheet in beikaobiao_wb.Worksheets:
        sheet_name = tmp_sheet.name
        sheet_name = sheet_name.encode("gbk")
        if "������" in sheet_name:
            beikaobiao_sheet = tmp_sheet
            break
    beikaobiao_pdf_path = output_dir + "/sdgkihgeisghiseghesgh.pdf"
    beikaobiao_sheet.ExportAsFixedFormat(0, beikaobiao_pdf_path)
    beikaobiao_pdf_file = open(beikaobiao_pdf_path, "rb")
    beikaobiao_pdf = PdfFileReader(beikaobiao_pdf_file)
    beikaobiao_page_num = beikaobiao_pdf.getNumPages()
    if beikaobiao_page_num < len(sheets.Worksheets):
        info.DisplayInfo("������ҳ��[{}] �� Ŀ¼ҳ��[{}] ��һ�£�".format(beikaobiao_page_num, len(sheets.Worksheets)))
        sheets.Close(False)
        beikaobiao_wb.Close(False)
        beikaobiao_pdf_file.close()
        os.remove(beikaobiao_pdf_path)
        excel.Application.Quit()
        return 1

    pdf_list = os.listdir(pdf_dir)

    for tmp_sheet in sheets.Worksheets:
        sheet_name = tmp_sheet.name
        sheet_name = sheet_name.encode("gbk")
        info.DisplayInfo("��ʼ����ҳ:" + sheet_name)
        pdf_fname = common.FindFNameByIdx(pdf_list, int(sheet_name))
        if not pdf_fname:
            info.DisplayInfo("û���ҵ��� " + sheet_name + " ҳ��Ӧ��pdf�ļ�")
            sheets.Close(False)
            beikaobiao_wb.Close(False)
            beikaobiao_pdf_file.close()
            os.remove(beikaobiao_pdf_path)
            excel.Application.Quit()
            return 1
        pdfPath = os.path.join(pdf_dir, pdf_fname)
        if (not os.path.exists(pdfPath)):
            info.DisplayInfo("pdf�ļ������ڣ�·����" + pdfPath)
            sheets.Close(False)
            beikaobiao_wb.Close(False)
            beikaobiao_pdf_file.close()
            os.remove(beikaobiao_pdf_path)
            excel.Application.Quit()
            return 1
        inputPDF = PdfFileReader(open(pdfPath, "rb"))
        catalog_pdf_path = output_dir + "/" + sheet_name + "kdjhdilrjhgdkrhn.pdf"
        tmp_sheet.ExportAsFixedFormat(0, catalog_pdf_path)
        catalog_pdf_file = open(catalog_pdf_path, "rb")
        catalog_pdf = PdfFileReader(catalog_pdf_file)

        output = PdfFileWriter()
        for page in catalog_pdf.pages:
            output.addPage(page)
        for page in inputPDF.pages:
            output.addPage(page)
        output.addPage(beikaobiao_pdf.getPage(int(sheet_name) - 1))
        output.write(open(output_dir + "/" + sheet_name + ".pdf", "wb"))
        catalog_pdf_file.close()
        os.remove(catalog_pdf_path)
    sheets.Close(False)
    beikaobiao_wb.Close(False)
    beikaobiao_pdf_file.close()
    os.remove(beikaobiao_pdf_path)
    excel.Application.Quit()
    return 0

def AutoSplitMain():
    warnings.simplefilter("ignore")
    info.DisplayInfo( "��ʼת��")

    excelFName = sys.argv[1]
    info.DisplayInfo("execel : " + excelFName)
    beikaobiaoFName = sys.argv[2]
    info.DisplayInfo("execel : " + beikaobiaoFName)
    pdfDir = sys.argv[3]
    info.DisplayInfo("pdf dir : " + pdfDir)
    outputDir = sys.argv[4]
    info.DisplayInfo("output dir : " + outputDir)
    appName = sys.argv[5]
    info.DisplayInfo("appName : " + appName)

    if Split(excelFName, beikaobiaoFName, pdfDir, outputDir, appName) == 0:
        info.DisplayInfo("�������")
        return

    info.DisplayInfo("����ʧ�ܣ��������ʾ�������������")

if __name__ == "__main__":
    common.SafeRunProgram(AutoSplitMain)