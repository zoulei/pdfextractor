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

# def getpageheader(img):
#     img.save(open("C:/Users/34695/Desktop/ggggggg.png", "wb"))
#     w,h = img.size
#     headlen = max(w,h)*0.08711
#     crop_rectangle = (0,0,w,headlen)
#     cropped_im = img.crop(crop_rectangle)
#     cropped_im.save(open("C:/Users/34695/Desktop/aaaaaaaaaup.png", "wb"))
#     try:
#         print pytesseract.image_to_string(cropped_im, lang="chi_sim", output_type=pytesseract.Output.STRING)
#         print pytesseract.image_to_osd(cropped_im, lang="chi_sim", output_type=pytesseract.Output.STRING)
#     except:
#         print "fail"
#     crop_rectangle = (0, h - headlen, w, h)
#     cropped_im = img.crop(crop_rectangle)
#     cropped_im.save(open("C:/Users/34695/Desktop/aaaaaaaaadown.png", "wb"))
#     try:
#         print pytesseract.image_to_string(cropped_im, lang="chi_sim", output_type=pytesseract.Output.STRING)
#         print pytesseract.image_to_osd(cropped_im, lang="chi_sim", output_type=pytesseract.Output.STRING)
#     except:
#         print "fail"
#     crop_rectangle = (0, 0, headlen, h)
#     cropped_im = img.crop(crop_rectangle)
#     cropped_im.save(open("C:/Users/34695/Desktop/aaaaaaaaaupleft.png", "wb"))
#     try:
#         print pytesseract.image_to_string(cropped_im, lang="chi_sim", output_type=pytesseract.Output.STRING)
#         print pytesseract.image_to_osd(cropped_im, lang="chi_sim", output_type=pytesseract.Output.STRING)
#     except:
#         print "fail"
#     crop_rectangle = (w-headlen, 0, w, h)
#     cropped_im = img.crop(crop_rectangle)
#     cropped_im.save(open("C:/Users/34695/Desktop/aaaaaaaaaright.png", "wb"))
#     try:
#         print pytesseract.image_to_string(cropped_im, lang="chi_sim", output_type=pytesseract.Output.STRING)
#         print pytesseract.image_to_osd(cropped_im, lang="chi_sim", output_type=pytesseract.Output.STRING)
#
#     except:
#         print "fail"


def modifybluecover(img):
    # img.save(open("C:/Users/34695/Desktop/ggggggg.png", "wb"))
    img = img.convert("RGBA")
    data = np.array(img)
    red, green, blue, alpha = data.T
    grey_areas = (red >= 170) & (red <= 210) & (green >= 220) & (green <= 252) & (blue >= 220) & (blue <= 251)
    data[..., :-1][grey_areas.T] = (255, 255, 255)
    img = Image.fromarray(data)
    # img.save(open("C:/Users/34695/Desktop/ffffff.png","wb"))
    return img

def modifyimage(img):
    img = img.convert("RGBA")
    data = np.array(img)
    red,green,blue,alpha = data.T

    grey_areas = (red <= 50) | (green <= 50) | (blue <= 50)
    data[..., :-1][grey_areas.T] = (0, 0, 0)
    white_areas = (red > 50) & (green > 50) & (blue > 50)
    data[..., :-1][white_areas.T] = (255, 255, 255)
    img = Image.fromarray(data)
    # img.save(open("C:/Users/34695/Desktop/nnnnnn.png", "wb"))
    return img

def ocrpage(arg):
    fname,idx = arg
    dpilist = [(100,0),(100,1), (200,0),(200,1)]
    # dpilist = [(100, 0), (100, 1), (200, 0), (200, 1), (500, 0), (500, 1), (750, 0), (750, 1), (1000, 0), (1000, 1)]
    for dpi, cimg in dpilist:
        try:
            images = convert_from_path(fname, dpi=dpi, first_page=idx + 1, last_page=idx + 1)
            if cimg == 1:
                images[0] = modifyimage(images[0])
            else:
                images[0] = modifybluecover(images[0])
            osdresult = pytesseract.image_to_osd(images[0], lang="chi_sim", output_type=pytesseract.Output.DICT)
            # print idx+1, (osdresult["orientation_conf"], osdresult["rotate"], dpi, cimg)
            if osdresult["orientation_conf"] < 1:
                continue
            return (osdresult["orientation_conf"],osdresult["rotate"], dpi)
        except:
            pass
            # traceback.print_exc()
    images = convert_from_path(fname, dpi=200, first_page=idx + 1, last_page=idx + 1)
    w, h = images[0].size
    if w > h:
        # print idx + 1,(-1, 270, -1)
        return (-1, 270, -1)
    else:
        # print idx + 1, (-1, 0, -1)
        return (-1, 0, -1)

def pdfdirection(fname, ofname):
    pdf_in = open(fname, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_in)
    pdf_writer = PyPDF2.PdfFileWriter()
    start = time.time()
    pool = multiprocessing.Pool(processes=multiprocessing.cpu_count())
    # ocrresult = [ocrpage((fname, 0)),]
    # return
    ocrresult = pool.map(ocrpage,[(fname, v) for v in xrange(pdf_reader.numPages)])
    # ocrresult = pool.map(ocrpage, [(fname, v) for v in xrange(1,2)])
    # print "multiprocess:", time.time() - start
    start = time.time()
    pool.close()
    foundfalse = False
    for idx in xrange(len(ocrresult)):
        page = pdf_reader.getPage(idx)
        conf, rotate, dpi = ocrresult[idx]
        if conf == -1:
            foundfalse = True
            error.RotateError(ofname,idx+1)
        if rotate != 0:
            page.rotateClockwise(rotate)
        pdf_writer.addPage(page)

    pdf_out = open(ofname, 'wb')
    pdf_writer.write(pdf_out)
    pdf_out.close()
    pdf_in.close()
    # print "write:", time.time() - start
    return foundfalse

def correctpdfdirection(sourcedir):
    sourcedir = common.Path(sourcedir)
    os.environ["PATH"] += ";"+os.getcwd()+"/dist/poppler-0.68.0/bin"
    os.environ["PATH"] += ";"+os.getcwd()+"/dist/Tesseract-OCR"
    info.DisplayInfo("检测文件")
    sourcedir = os.path.abspath(sourcedir)
    info.DisplayInfo(sourcedir)
    sourcedirlen = len(sourcedir)
    allfname = pathOperator.listallfiler(sourcedir)
    allfname = [v[sourcedirlen:] for v in allfname]
    originlen = len(allfname)
    rotatedir = sourcedir + "_rotate"
    rotatedir = os.path.abspath(rotatedir)
    allrotatedfname = pathOperator.listallfiler(rotatedir)
    rotatedirlen = len(rotatedir)
    allrotatedfname = [v[rotatedirlen:] for v in allrotatedfname]
    rotatelen = len(allrotatedfname)
    for v in allrotatedfname:
        if v in allfname:
            allfname.remove(v)

    for v in allfname:
        curfname = rotatedir + v
        curdirname = os.path.dirname(curfname)
        if not os.path.exists(curdirname):
            os.makedirs(curdirname)

    info.DisplayInfo("开始转换文件，处理结果存储在如下路径中：\n" + rotatedir.encode("gbk"))
    for idx,v in enumerate(allfname):
        foundfalse = pdfdirection(sourcedir+"/"+v,rotatedir+"/"+v)
        if foundfalse:
            info.DisplayInfo("完成文件" + (rotatedir + "/" + v).encode("gbk") + "，本文档中存在一些无法识别的页面，需要人工处理，相关信息存储在‘错误.txt’文件中")
        else:
            info.DisplayInfo("完成文件" + (rotatedir+"/"+v).encode("gbk") + "")
        info.DisplayInfo("已完成第 " + str(idx+rotatelen + 1).encode("gbk") + " 个pdf文件，共 " + str(originlen).encode("gbk") + " 个pdf文件")

    info.DisplayInfo("全部完成")
    return rotatedir

def correctpdfmain():
    error.CreateLog("错误.txt")
    excelFName = sys.argv[1]
    # excelFName = "C:/Users/34695/Desktop/testdir1"

    rotatedir = correctpdfdirection(excelFName)
    info.DisplayInfo("处理完毕,结果存储在如下路径中：\n"+ rotatedir.encode("gbk"))
    error.Finish()

# def testfunc(arg):
#     print arg

if __name__ == "__main__":
    multiprocessing.freeze_support()
    # pool = multiprocessing.Pool(processes=3)
    # pool.map(testfunc,range(2))
    common.SafeRunProgram(correctpdfmain)

# print 123
# time.sleep(15)
# pdfdirection("C:\Users/34695\Desktop/02.pdf","C:\Users/34695\Desktop/02ggg.pdf")