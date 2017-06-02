# -*- coding:gbk -*-
from distutils.core import setup
import common
import py2exe
setup(console=["autoSplit.py","deleteABCD.py","mercolumn.py","gencover.py","checksplitresult.py","packfile.py"])

# common.RenameFile("dist/autoSplit.exe","dist/自动提取pdf.exe")
# common.RenameFile("dist/deleteABCD.exe","dist/删除ABCD.exe")
#common.RenameFile("dist/mercolumn.exe","dist/合并空格.exe")
