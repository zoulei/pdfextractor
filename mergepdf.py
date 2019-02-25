# -*- coding:gbk -*-
import pytesseract
from pdf2image import convert_from_path, convert_from_bytes
import time
from PIL import Image
import PyPDF2
import os
import multiprocessing
import pathOperator
import info
import common
import error
import sys
import numpy as np
import traceback
import time

Image.MAX_IMAGE_PIXELS = None

def executemerge(sourcedirname, targetdirname, prefix, partdata):
    merger = PyPDF2.PdfFileMerger()
    for partidx in partdata:
        merger.append(PyPDF2.PdfFileReader(file(sourcedirname+prefix+"_"+partidx+".pdf","rb")))
    dirname = os.path.dirname(targetdirname +  prefix)
    if not os.path.exists(dirname):
        os.makedirs(dirname)
    merger.write(targetdirname+prefix+".pdf")

def mergepdf(sourcedir):
    sourcedir = common.Path(sourcedir)
    # os.environ["PATH"] += ";" + os.getcwd() + "/dist/poppler-0.68.0/bin"
    # os.environ["PATH"] += ";" + os.getcwd() + "/dist/Tesseract-OCR"
    info.DisplayInfo("检测文件")
    sourcedir = os.path.abspath(sourcedir)
    info.DisplayInfo(sourcedir)
    sourcedirlen = len(sourcedir)
    allfname = pathOperator.listallfiler(sourcedir)
    allfname = [v[sourcedirlen:] for v in allfname]
    allfname = [v for v in allfname if v.endswith(".pdf")]
    allfname.sort()
    mergedict = {}
    for fname in allfname:
        dotpos = fname.rfind(".")
        dashpos = fname.rfind("_")
        if dotpos == -1 or dashpos == -1:
            continue
        prefixfname = fname[:dashpos]
        try:
            partidx = int(fname[dashpos+1:dotpos])
        except:
            continue
        if prefixfname not in mergedict:
            mergedict[prefixfname] = []
        mergedict[prefixfname].append(fname[dashpos+1:dotpos])
    rotatedir = sourcedir + "_merge"
    for fnameprefix, partdata in mergedict.items():
        if len(partdata) <= 1:
            continue
        partdata.sort(key=lambda v:int(v))
        executemerge(sourcedir, rotatedir, fnameprefix, partdata)
    return rotatedir

def mergepdfmain():
    error.CreateLog("错误.txt")
    excelFName = sys.argv[1]
    # excelFName = "C:/Users/34695/Desktop/testdir1"

    rotatedir = mergepdf(excelFName)
    info.DisplayInfo("处理完毕,结果存储在如下路径中：\n" + rotatedir)
    error.RealFinish()

if __name__ == "__main__":
    common.SafeRunProgram(mergepdfmain)
