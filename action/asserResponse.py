from utils.log.excuteWriteLog import *
class asserResponse(object):
    @classmethod
    def checkRsult(cls,response,checkPoint):
        """
        :param response:
        :param checkPoint:
        :return:
        """
        excutelog("info","excute asser Response starting")
        rightNum=0
        wrongNum=0
        failedReason = {}
        if checkPoint!=None and response!=None:
            for key,value in eval(checkPoint).items():
                if key in list(eval(response).keys()):
                    if eval(response)[key]!=value:
                        failedReason[key]="预期：%s--->实际：%s"%(value,eval(response)[key])
                        wrongNum+=1
                    else:
                        rightNum+=1
                else:
                    failedReason[key] = "预期：%s--->实际：the key not in response" % value
                    wrongNum += 1
            if len(eval(checkPoint)) == rightNum:
                excutelog("info","asser response success,the case pass")
                return ['True']
            else:
                excutelog("info", "asser response success,the case fail ")
                excutelog("info", "[False, 断言失败次数：%s, 失败原因：%s]"%(wrongNum,failedReason))
                return ['False', '断言失败次数：%s' % wrongNum, '失败原因：%s' % failedReason]
        elif response==None:
            excutelog("info", "asser response success,the case fail ")
            excutelog("info", "[False, 断言失败次数：1, 失败原因：response is None]")
            return ['False', '断言失败次数：1', '失败原因：response is None']
        else:
            return ['True']



if __name__=="__main__":
    print(asserResponse.checkRsult("{'login_name':'wqliu','user_id':'0001'}","{'login_name':'wqliu','user_id':'0001'}"))


