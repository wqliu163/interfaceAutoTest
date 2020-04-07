import requests
import json
from utils.log.excuteWriteLog import *

class httpRequest(object):

    @classmethod
    def httpRequest(cls,requestUrl,requestMethod,paramsType,spacer,requestData=None,header=None):
        """
        :param requestUrl:        (str)     请求url
        :param requestMethod:    (str)      请求方法，get，post
        :param paramsType:       (str)      请求参数类型，如url，form（body）
        :param spacer:          (str)       url与参数间隔符，如 ? /
        :param requestData:     (dict)      请求的参数
        :param requestHeaders:  (dict)      请求头
        :return:
        """
        if requestMethod=="get":
            if paramsType=="url" and isinstance(requestData,dict):
                parameter=""
                for key,value in requestData.items():
                    parameter+=key+"="+value+"&"
                parameter=parameter.strip("&")
                print(parameter)
                completeUrl="%s%s%s"%(requestUrl,spacer,parameter)
                response=requests.get(completeUrl,headers=header)
                excutelog("info","request url is %s"%completeUrl)
                excutelog("info","response=%s"%response)
                print(response)
                return response
            elif paramsType=="params" and isinstance(requestData,dict):
                parameter=json.dumps(requestData)
                response=requests.get(requestUrl,params=parameter,headers=header)
                excutelog("info", "response=%s" % response)
                return response
            else:
                response = requests.request("get",requestUrl, params=requestData, headers=header)
                excutelog("info", "response=%s" % response)
                return response
        elif paramsType=="post":
            if paramsType=="form" and isinstance(requestData,dict):
                parameter=json.dumps(requestData)
                response=requests.post(requestUrl,data=parameter,headers=header)
                excutelog("info", "response=%s" % response)
                return response
            elif paramsType=="json" and isinstance(requestData,dict):
                response = requests.post(requestUrl, json=requestData, headers=header)
                excutelog("info", "response=%s" % response)
                return response
            else:
                print("requestData must be dict type !")
                excutelog("info","requestData must be dict type !")

if __name__=="__main__":
    httpRequest.httpRequest("https://fanyi.baidu.com/translate","get","url","?",requestData={"aldtype":"16047","query":"","keyfrom":"baidu","smartresult":"dict","lang":"auto2zh"},header={"User-Agent": "Chrome/68.0.3440.106"})


