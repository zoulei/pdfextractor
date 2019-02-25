# -*- coding:gbk -*-
# from cx_Freeze import setup, Executable
from distutils.core import setup
import common
import py2exe

# setup(name="correctpdfdirection",
#       version="0.1",
#       description="",
#       executables=[Executable("correctpdfdirection.py")],
#       options={
#             'build_exe': {
#                   'includes': ['atexit', 'numpy.core._methods', 'numpy.lib.format'],
#             }
#       }
#       )

# how to use :
# under cmd : python setup.py py2exe
setup(console=["autoSplit.py","deleteABCD.py","mercolumn.py","gencover.py","checksplitresult.py","packfile.py","newgencover.py","autofillxlsfile.py","correctpdfdirection.py"])

# common.RenameFile("dist/autoSplit.exe","dist/自动提取pdf.exe")
# common.RenameFile("dist/deleteABCD.exe","dist/删除ABCD.exe")
#common.RenameFile("dist/mercolumn.exe","dist/合并空格.exe")
