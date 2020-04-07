from utils.handleExcel import *
from config.config import *
from action.getRely import *
from action.asserResponse import *
from utils.log.excuteWriteLog import *
from utils.httpRequest import *
from utils.sendEmail import *
from utils.getDate import *
import traceback

def main():
    excutelog("info","starting excute script testing interface API ")
    excuteCasePassNum=0
    excuteCaseFailNum = 0
    excutelog("info", "开始遍历需要执行的API接口")
    for index,cell in enumerate(handleExcel.getColumnsObject(getRootPath()+"\\data\\case.xlsx","API",columnNo=API_active)[1:],2):
        #print(type(index),cell.value)
        if cell.value=="Y" or cell.value=="y":
            apiName=handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx","API",columnNo=ord(API_apiName)-64,rowNo=index)
            requestUrl=handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx","API",columnNo=ord(API_requestUrl)-64,rowNo=index)
            requestMethod=handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx","API",columnNo=ord(API_requestMothod)-64,rowNo=index)
            parameterType=handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx","API",columnNo=ord(API_paramsType)-64,rowNo=index)
            apiCaseSheetName=handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx","API",columnNo=ord(API_apiTestCaseFileName)-64,rowNo=index)
            excutelog("info","requestUrl=%s"%requestUrl)
            excutelog("info", "requestMethod=%s" %requestMethod)
            excutelog("info", "parameterType=%s" %parameterType)
            excutelog("info", "apiCaseSheetName=%s" %apiCaseSheetName)
            #print(apiName,requestUrl,requestMethod,parameterType,apiCaseSheetName)
            excutelog("info", "开始遍历需要执行的API接口用例")
            for idx,cll in enumerate(handleExcel.getColumnsObject(getRootPath()+"\\data\\case.xlsx",apiCaseSheetName,columnNo=CASE_active)[1:],2):
                if cll.value=="Y" or cll.value=="y":
                    requestData=handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx",apiCaseSheetName,columnNo=ord(CASE_requestData)-64,rowNo=idx)
                    relyData=handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx",apiCaseSheetName,columnNo=ord(CASE_relyData)-64,rowNo=idx)
                    checkPoint=handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx",apiCaseSheetName,columnNo=ord(CASE_checkPoint)-64,rowNo=idx)
                    excutelog("info","requestData=%s /n relyData=%s  /n checkPoint=%s"%(requestData,relyData,checkPoint))
                    #print(requestData)
                    requestData_new=eval(requestData)
                    #print("requestData_new",requestData_new)
                    for k,w in getRely.GetRely(requestData,relyData).items():
                        requestData_new[k]=w
                    #print(requestData_new)
                    excutelog("info","requestData_new=%s" %requestData_new)
                    excutelog("info", "relyData=%s" %relyData)
                    excutelog("info", "checkPoint=%s" % checkPoint)
                    response=httpRequest.httpRequest(requestUrl,requestMethod,parameterType,"\\",requestData=requestData_new)
                    #print("55555555555",response)
                    asserResult=asserResponse.checkRsult(response,checkPoint)
                    excutelog("info","asserResult=%s"%asserResult)
                    #print("6666666%s"%asserResult)
                    if len(asserResult)>1:
                        excuteCaseFailNum+=1
                        handleExcel.writeValueInCell(getRootPath()+"\\data\\case.xlsx",apiCaseSheetName,"fail",rowNo=idx,columnNo=ord(CASE_status)-64,colour="red")
                        if response!=None:
                            handleExcel.writeValueInCell(getRootPath() + "\\data\\case.xlsx",apiCaseSheetName, response.status_code, rowNo=idx,columnNo=ord(CASE_responseCode)-64, colour="black")
                            handleExcel.writeValueInCell(getRootPath() + "\\data\\case.xlsx",apiCaseSheetName, response.text,rowNo=idx, columnNo=ord(CASE_responseData)-64,colour="black")
                        else:
                            handleExcel.writeValueInCell(getRootPath() + "\\data\\case.xlsx", apiCaseSheetName,"404", rowNo=idx,columnNo=ord(CASE_responseCode) - 64, colour="black")
                            handleExcel.writeValueInCell(getRootPath() + "\\data\\case.xlsx", apiCaseSheetName,"None", rowNo=idx, columnNo=ord(CASE_responseData) - 64,colour="black")
                        handleExcel.writeValueInCell(getRootPath() + "\\data\\case.xlsx",apiCaseSheetName, asserResult[2], rowNo=idx,columnNo=ord(CASE_failedReason)-64, colour="red")
                        handleExcel.writeValueInCell(getRootPath() + "\\data\\case.xlsx", apiCaseSheetName,getDate.getDate() + "  " + getDate.getTime(), rowNo=idx,columnNo=ord(CASE_testTime) - 64, colour="red")
                    else:
                        excuteCasePassNum+=1
                        handleExcel.writeValueInCell(getRootPath() + "\\data\\case.xlsx",apiCaseSheetName, "pass", rowNo=idx,columnNo=ord(CASE_status)-64, colour="black")
                        handleExcel.writeValueInCell(getRootPath() + "\\data\\case.xlsx",apiCaseSheetName, response.status_code,rowNo=idx, columnNo=ord(CASE_responseCode)-64,colour="black")
                        handleExcel.writeValueInCell(getRootPath() + "\\data\\case.xlsx",apiCaseSheetName, response.text, rowNo=idx,columnNo=ord(CASE_responseData)-64, colour="black")
                        handleExcel.writeValueInCell(getRootPath() + "\\data\\case.xlsx",apiCaseSheetName, getDate.getDate()+"  "+getDate.getTime(),rowNo=idx, columnNo=ord(CASE_testTime)-64,colour="red")

    if switch==1:
        sendEmail.sendEmailWithCarry("接口测试结果","执行成功用例数%s,执行失败用例数%s"%(excuteCasePassNum,excuteCaseFailNum),toEmails,filesPath=getRootPath() + "\\data\\case.xlsx")
    else:
        print("don't send email !")
if __name__=="__main__":
    main()