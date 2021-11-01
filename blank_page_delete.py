# -*- coding:gbk -*-
from pyPdf import PdfFileWriter, PdfFileReader
import os.path
import traceback
import error
import info
import common
import sys
from multiprocessing import Pool
from multiprocessing import cpu_count


import fitz
import io
from PIL import Image

import shutil

WHITE_SIZE_THRE = 0.995


def is_blank_page(i, source_path):
    pdf_file = fitz.open(source_path)
    page = pdf_file[i]
    image_list = page.getImageList()
    # print "ppp:", len(image_list)
    # import pprint
    # pp = pprint.PrettyPrinter(indent=4)
    white_point = 0
    total_point = 0
    for img in image_list:
        # get the XREF of the image
        xref = img[0]
        # extract the image bytes
        base_image = pdf_file.extractImage(xref)
        # pp.pprint(base_image)
        image_bytes = base_image["image"]
        image = Image.open(io.BytesIO(image_bytes))
        # image = image.convert("L")
        # image.save(open("tmp.png", "wb"))
        for x in range(image.width):
            for y in range(image.height):
                pixel = image.getpixel((x, y))
                if pixel[0] > 240 and pixel[1] > 240 and pixel[2] > 240:
                    white_point += 1
                # print "ttttt:", image.getpixel
                # exit(0)
                # if image.getpixel((x, y)) == 255:
                #     white_point += 1
        total_point += image.width * image.height
    if white_point * 1.0 / total_point > WHITE_SIZE_THRE:
        return [True, white_point * 1.0 / total_point]
    return [False, white_point * 1.0 / total_point]

class DeleteBlankPage:
    def __init__(self, source_data_file, delete_label, start_page=0, end_page=0):
        self.m_data_file = os.path.basename(source_data_file)
        self.m_source_data_dir = os.path.dirname(source_data_file)
        self.m_target_data_dir = self.m_source_data_dir + "___"
        if not os.path.exists(self.m_target_data_dir):
            os.mkdir(self.m_target_data_dir.decode("gbk"))
        self.m_start_page = start_page
        self.m_end_page = end_page
        self.m_delete_label = delete_label
        # print "data_file:", self.m_data_file
        # print "source_data_dir:", self.m_source_data_dir
        # print "target_data_dir:", self.m_target_data_dir
        # self.m_source_data_dir = common.Path(source_data_dir)
        # while self.m_source_data_dir.endswith("/") or self.m_source_data_dir.endswith("\\"):
        #     self.m_source_data_dir = self.m_source_data_dir[:-1]
        # self.m_target_data_dir = self.m_source_data_dir.decode("gbk") + "_无空白页".decode("gbk")
        # if os.path.exists(self.m_target_data_dir):
        #     shutil.rmtree(self.m_target_data_dir)
        # os.mkdir(self.m_target_data_dir)

    def run(self):
        self.process_file(self.m_source_data_dir + "/" + self.m_data_file, self.m_target_data_dir + "/" + self.m_data_file)

        # path_stack = os.listdir(self.m_source_data_dir)
        # while len(path_stack):
        #     current_file_name = path_stack.pop()
        #     current_full_path = self.m_source_data_dir + "/" + current_file_name
        #     if os.path.isdir(current_full_path):
        #         to_add_path = [current_file_name + "/" + v for v in os.listdir(current_full_path)]
        #         path_stack.extend(to_add_path)
        #     elif current_file_name.endswith(".pdf"):
        #         self.process_file(current_file_name)

    def tongji_page_grey(self, image):
        width_dis = [0] * 256
        for x in range(image.width):
            for y in range(image.height):
                width_dis[image.getpixel((x, y))] += 1
        for i in range(1, len(width_dis)):
            width_dis[i] += width_dis[i - 1]
        for i, v in enumerate(width_dis):
            print i, v
        print "pcg:", width_dis[254] * 1.0 / width_dis[255]



    def process_file(self, file_name, target_file_name):
        target_file_name = target_file_name.decode("gbk")
        if os.path.exists(target_file_name):
            os.remove(target_file_name)
        source_path = file_name.decode("gbk")
        inputPDF = PdfFileReader(open(source_path, "rb"))
        output = PdfFileWriter()
        page_size = inputPDF.getNumPages()
        end_page = self.m_end_page
        if end_page <= 0:
            end_page = page_size
        for i in range(0, self.m_start_page):
            output.addPage(inputPDF.getPage(i))
        for i in range(self.m_start_page, end_page):
            if i % 2 == self.m_delete_label % 2:
                output.addPage(inputPDF.getPage(i))
        for i in range(end_page, page_size):
            output.addPage(inputPDF.getPage(i))
        # if True:
        #     pool = Pool(cpu_count())
        #     task_list = list()
        #     for i in range(page_size):
        #         task = pool.apply_async(is_blank_page, args=[i, source_path])
        #         task_list.append(task)
        #     write_size = 0
        #     for i, task in enumerate(task_list):
        #         blank, rate = task.get()
        #
        #         # print i, blank, rate
        #         # if i >= 164 and i % 2 == 0 and not blank:
        #         #     print "====================================================================="
        #
        #         if not blank:
        #             output.addPage(inputPDF.getPage(i))
        #             write_size += 1
        #             # print "add to:", write_size
        # else:
        #     # for i in range(page_size):
        #     for i in range(0, 1):
        #         blank, rate = is_blank_page(i, source_path)
        #         print "tttttttttttt:", blank, rate
        #         if not blank:
        #             output.addPage(inputPDF.getPage(i))
        output.write(open(target_file_name, "wb"))

def delete_blank_page_main():
    # info.DisplayInfo( "请输入excel文件路径.")
    error.CreateLog("错误.txt")
    source_data_dir = common.Path(sys.argv[1])
    info.DisplayInfo( "开始处理文件")
    delete_label = int(sys.argv[2])
    start_page = 0
    end_page = 0
    if len(sys.argv) > 3:
        start_page = max(int(sys.argv[3]) - 1, 0)
    if len(sys.argv) > 4:
        end_page = int(sys.argv[4])
    DeleteBlankPage(source_data_dir, delete_label, start_page, end_page).run()
    info.DisplayInfo( "处理文件结束")

if __name__ == "__main__":
    common.SafeRunProgram(delete_blank_page_main)