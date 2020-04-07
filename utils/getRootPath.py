import os

def getRootPath(filePath=__file__):
    rootPath=os.path.dirname(os.path.dirname(filePath).replace("/","\\"))
    #print(type(rootPath))
    #print(rootPath)
    return rootPath

if __name__=="__main__":
    #getRootPath(r"E:\\whiteMouseProduct\UIAutoTestWithSelenium")
    getRootPath()