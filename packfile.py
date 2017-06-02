# -*- coding: gbk -*-
import os
import os.path
import common
import sys
import info

# this function generater directory for each file and drag the files into their directory
def createFile(dir):
	fileList = os.listdir(dir)
	for filename in fileList:
		dotPos = filename.find(".")
		subdir = os.path.join(dir,filename[:dotPos])
		fullfname = os.path.join(dir,filename)
		if os.path.isdir(fullfname):
                        continue
		newfname = os.path.join(subdir,filename)
		if not os.path.exists(subdir):
			os.mkdir(subdir)
		if os.path.exists(newfname):
			os.remove(newfname)
		os.rename(fullfname,newfname)

def AutoPackFile():
	dirpath = sys.argv[1]
	dirpath = common.Path(dirpath)
	info.DisplayInfo("开始拖文件")
	createFile(dirpath)
	info.DisplayInfo("拖文件完毕")

if __name__ == "__main__":
    common.SafeRunProgram(AutoPackFile)
	# dir = raw_input("请输入文件目录")
	# createFile(dir)
