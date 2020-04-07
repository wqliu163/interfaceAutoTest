from utils.log.excuteWriteLog import *
from utils.handleExcel import *
from utils.getRootPath import *
from config.config import *

class getRely(object):
    @classmethod
    def GetRely(cls,requestData, relyData):
        """
        params：requestData    （str，dict）  请求参数
        params：relyData       （str，dict）  依赖关系，格式{请求或者响应：{依赖参数的key：接口名->用例编号id}}
        result：requestData    （dict）       根据依赖关系处理后的请求参数
        """
        if not requestData:  # 如果请求参数为空，则不需要处理依赖
            return
        elif requestData and not relyData:  # 如果没有依赖（关联），直接返回请求参数
            if isinstance(requestData, str):
                print(requestData)
                print(type(requestData))
                return eval(requestData)
            elif isinstance(requestData, dict):
                return requestData
        else:  # 如果存在依赖（关联）
            if isinstance(requestData, str):
                requestData = eval(requestData)
            if isinstance(relyData, str):
                relyData = eval(relyData)
            for key, value in relyData.items():  # {"request":{"username->username1":"register->1","password->password1":"register->1"}}
                if key == "request":
                    excutelog("info","-----------关联请求参数----------")
                    for k, v in value.items():
                        relyKey, requestParamsKey = k.split("->")
                        interfaceName, caseId = v.split("->")
                        excutelog("info","上一个请求中关联请求参数key：-------%s" % relyKey)
                        excutelog("info","当前的请求参数中需要关联的key：-------%s" % requestParamsKey)
                        excutelog("info","被关联的接口名字：-------%s" % interfaceName)
                        excutelog("info","被关联接口用例的序号：-------%s" % caseId)
                        print('上一个请求中关联的参数key', relyKey)
                        print('当前的请求参数中需要关联的key:', requestParamsKey)
                        print('interfaceName:', interfaceName)
                        print('caseId:', caseId)
                        # 遍历API的接口名一列，根据接口名找到对应case用例的sheet名
                        for idx, vle in enumerate(handleExcel.getColumnsObject(getRootPath()+"\\data\\case.xlsx","API",API_apiName)[1:], 2):
                            print('接口序号：', idx)
                            print('接口名称：', vle.value)
                            if vle.value == interfaceName:
                                print(idx, API_apiTestCaseFileName)
                                apiCaseSheet = handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx","API",columnNo=ord(API_apiTestCaseFileName)-64,rowNo=idx)  # 依赖的接口用例sheet
                                excutelog("info","被关联用例所在sheet表：-------%s" % apiCaseSheet)
                                print(apiCaseSheet)
                                val = eval(handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx",apiCaseSheet,columnNo=ord(CASE_requestData)-64,rowNo=int(caseId) + 1))[relyKey]
                                requestData[requestParamsKey] = val
                    print(requestData)
                    excutelog("info","处理完依赖关系的请求参数：-------%s" % requestData)
                    return requestData
                elif key == "response":
                    excutelog("info","-----------关联响应body----------")
                    for k, v in value.items():
                        print(k, v)
                        interfaceName, caseId = v.split("->")
                        relyKey, requestParamsKey = k.split("->")
                        excutelog("info","上一个请求响应中关联请求参数key：-------%s" % relyKey)
                        excutelog("info","当前的请求参数中需要关联的key：-------%s" % requestParamsKey)
                        excutelog("info","被关联的接口名字：-------%s" % interfaceName)
                        excutelog("info","被关联接口用例的序号：-------%s" % caseId)
                        print('响应body中的关联参数key:', relyKey)
                        print('当前的请求参数中需要关联的key:', requestParamsKey)
                        print('接口序号：', caseId)
                        print('接口名称：', interfaceName)
                        for idx, vle in enumerate(handleExcel.getColumnsObject(getRootPath()+"\\data\\case.xlsx","API", API_apiName)[1:], 2):
                            print(idx, vle)
                            if vle.value == interfaceName:
                                print(idx, API_apiTestCaseFileName)
                                apiCaseSheet = handleExcel.getValueOfCell(getRootPath()+"\\data\\case.xlsx","API",columnNo=ord(API_apiTestCaseFileName)-64,rowNo=idx)
                                excutelog("info","被关联用例所在sheet表：-------%s" % apiCaseSheet)
                                print('apiCaseSheet:', apiCaseSheet)
                                print(ord(CASE_responseData) - 64)
                                print(apiCaseSheet)
                                val = eval(handleExcel.getValueOfCell(apiCaseSheet, coordinate=None,columnNo=ord(CASE_responseData)-64,rowNo=int(caseId) + 1))[relyKey]
                                print('val:', val)
                                requestData[requestParamsKey] = val
                    print(requestData)
                    excutelog("info","处理完依赖关系的请求参数：-------%s" % requestData)
                    return requestData

if __name__=="__main__":
    pass