from utils.getDate import *
from utils.getRootPath import *
from utils.log.config.logConf import *

class weLog(object):
    levelList=["debug","info","error"]
    def __init__(self,SetLevel,logfilePath):
        self.logfilePath=logfilePath.strip("\\")
        if isinstance(SetLevel,int):
            if SetLevel>= 0 and SetLevel<=2:
                self.SetLevel=SetLevel
            else:
                self.SetLevel =1
        if isinstance(SetLevel, str):
            if SetLevel.lower() in self.levelList:
                self.SetLevel=self.levelList.index(SetLevel.lower())
            else:
                self.SetLevel = 1
        else:
            self.SetLevel = 1

    def writeLogsInFile(self,level,log_content):
        if os.path.exists(self.logfilePath):
            if os.path.exists(self.logfilePath + '\\' + getDate.getDate() + '.log'):
                file = open(self.logfilePath + '\\' + getDate.getDate() + '.log', "r", encoding='utf-8')
                line_content = file.readline()
                if line_content.strip() == '':
                    file.close()
                    with open(self.logfilePath + '\\' + getDate.getDate() + '.log', 'a+',
                              encoding='utf-8') as fq:
                        fq.write(getDate.getDate() + ' ' + getDate.getTime() + '   WELOG     '+level.upper()+"    ")
                        fq.write(log_content)
                else:
                    file.close()
                    with open(self.logfilePath + '\\' + getDate.getDate() + '.log', 'a+',
                              encoding='utf-8') as fq:
                        fq.write("\n")
                        fq.write(getDate.getDate() + ' ' + getDate.getTime() + '   WELOG     '+level.upper()+"    ")
                        fq.write(log_content)
            else:
                file = open(self.logfilePath + '\\' + getDate.getDate() + '.log', "w",
                            encoding='utf-8')  # 创建文件
                file.close()
                with open(self.logfilePath + '\\' + getDate.getDate() + '.log', 'r+',
                          encoding='utf-8') as fq:
                    fq.write(getDate.getDate() + ' ' + getDate.getTime() + '   WELOG     '+level.upper()+"    ")
                    fq.write(log_content)
        else:
            if os.path.exists(getRootPath()+'\\logs\\'+getDate.getDate()+'.log'):
                file = open(getRootPath() + '\\logs\\' + getDate.getDate() + '.log', "r",encoding='utf-8')
                line_content=file.readline()
                if line_content.strip() == '':
                    file.close()
                    with open(getRootPath()+'\\logs\\'+getDate.getDate()+'.log','a+',encoding='utf-8') as fq:
                        fq.write(getDate.getDate() + ' ' + getDate.getTime() + '   WELOG     '+level.upper()+"    ")
                        fq.write(log_content)
                else:
                    file.close()
                    with open(getRootPath() + '\\logs\\' + getDate.getDate() + '.log', 'a+', encoding='utf-8') as fq:
                        fq.write("\n")
                        fq.write(getDate.getDate() + ' ' + getDate.getTime() + '   WELOG     '+level.upper()+"    ")
                        fq.write(log_content)
            else:
                file=open(getRootPath()+'\\logs\\'+getDate.getDate()+'.log',"w",encoding='utf-8')        #创建文件
                file.close()
                with open(getRootPath()+'\\logs\\'+getDate.getDate()+'.log','r+',encoding='utf-8') as fq:
                    fq.write(getDate.getDate() + ' ' + getDate.getTime() + '   WELOG     '+level.upper()+"    ")
                    fq.write(log_content)

# if __name__=='__main__':
#     logHandler.logging("good")
#     logHandler.logging("test")