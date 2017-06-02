import os

# list all file and dir under path dir
def listall(path):
    fileList = os.listdir(path)
    dirName = os.path.dirname(path)
    return [os.path.join(path,v) for v in fileList]

# list all file under path dir recursively
def listallfiler(path):
    allFileList = []

    dirstack = [path]
    while dirstack:
        path = dirstack.pop()

        fileList = os.listdir(path)
        dirName = os.path.dirname(path)

        for fname in fileList:
            subfilepath = os.path.join(path,fname)
            if os.path.isdir(subfilepath):
                dirstack.append(subfilepath)
            else:
                allFileList.append(subfilepath)
    return allFileList


# list all file under path dir
def listfile(path):
    fileList = os.listdir(path)
    fullpathList = [os.path.join(path, v) for v in fileList]
    return [v for v in fullpathList if os.path.isfile(v)]

# list all dir under path dir
def listdir(path):
    fileList = os.listdir(path)
    fullpathList = [os.path.join(path, v) for v in fileList]
    return [v for v in fullpathList if os.path.isdir(v)]