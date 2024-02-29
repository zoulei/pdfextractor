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
# under cmd : c:\python27\python.exe setup.py py2exe
# setup(console=["autoSplit.py","deleteABCD.py","mercolumn.py","gencover.py","checksplitresult.py","packfile.py","newgencover.py","autofillxlsfile.py","correctpdfdirection.py","mergepdf.py","blank_page_delete.py"])
setup(console=["autoSplit_1.py", "add_catalog_beikaobiao.py", "add_catalog.py"], options={
        "py2exe":{
             "dll_excludes":["libopenblas.TXA6YQSD3GCQQC22GEQ54J2UDCXDXHWN.gfortran-win_amd64.dll",
                             "libdet.KNZJ5W323CIP452TKKVK72OCSS32QOKE.gfortran-win_amd64.dll",
                             "libdop853.6TJTQZW3I3Q3QIDQHEOBEZKJ3NYRXI4B.gfortran-win_amd64.dll",
                             "libchkder.G7WSOGIYYQO3UWFVEZ3PPXCXR53ADVPA.gfortran-win_amd64.dll",
                             "libdcosqb.YMN7XEXYADIEZSKAGEVNR4E3MD7AXDG2.gfortran-win_amd64.dll",
                             "libbispeu.7AH3PCQ2E2NGLC3AQD7FFAH73KGJTZCJ.gfortran-win_amd64.dll",
                             "MSVCP90.dll"],
             "compressed":1,
             "optimize":2,
             # "bundle_files":1
         }
     },)

# common.RenameFile("dist/autoSplit.exe","dist/自动提取pdf.exe")
# common.RenameFile("dist/deleteABCD.exe","dist/删除ABCD.exe")
#common.RenameFile("dist/mercolumn.exe","dist/合并空格.exe")
