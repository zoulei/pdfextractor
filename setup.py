# -*- coding:gbk -*-
from distutils.core import setup
import common
import py2exe
setup(console=["autoSplit.py","deleteABCD.py","mercolumn.py","gencover.py","checksplitresult.py","packfile.py"])

# common.RenameFile("dist/autoSplit.exe","dist/�Զ���ȡpdf.exe")
# common.RenameFile("dist/deleteABCD.exe","dist/ɾ��ABCD.exe")
#common.RenameFile("dist/mercolumn.exe","dist/�ϲ��ո�.exe")
