import time
class getDate(object):
    @classmethod
    def getDate(cls,spacer=None):
        if spacer==None:
            date=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())[:10]
            #print(type(date))
            #print(date[8:10])
            #print(date)
            return date
        else:
            date = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())[:10]
            date=date.replace("-",spacer)
            # print(type(date))
            # print(date[8:10])
            #print(date)
            return date

    @classmethod
    def getTime(cls):
        times = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())[11:]
        #print(times[:5].replace(':',''))
        #print(int(times.replace(':','')[:4]))
        #print(times)
        return times

if __name__=="__main__":
    getDate.getDate()
    getDate.getTime()