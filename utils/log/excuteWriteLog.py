from utils.log.weLog import *

def excutelog(level, logcontent):
    logwriter = weLog(SetLevel, logfilePath)
    if level.lower() in logwriter.levelList:
        if logwriter.levelList.index(level.lower()) >= logwriter.SetLevel:
            logwriter.writeLogsInFile(level,logcontent)
        else:
            pass
    else:
        print("input loglevel unknows ")

if __name__=="__main__":
    excutelog("debug","debug testing")
    excutelog("info", "this is whiteMouse write logs")
    excutelog("error", "input type error")