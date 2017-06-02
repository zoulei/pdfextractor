# -*- coding:gbk -*-

import os
import sys
import pathOperator
import common

def ClearAllTxt(dirname):
    txtfileList = pathOperator.listallfiler(dirname)
    for fname in txtfileList:
        # print fname
        if fname.endswith(".txt"):
            os.remove(fname)

if __name__ == "__main__":
    dirname = raw_input()
    ClearAllTxt(common.Path(dirname))