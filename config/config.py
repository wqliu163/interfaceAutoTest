
#邮箱配置
smtp_port = 465  # smtp服务器SSL端口号，默认是465
smtp_host = "smtp.qq.com"  # 发送邮件的smtp服务器（从QQ邮箱中取得）
smtp_from_email = "1065677354@qq.com"  # 用于登录smtp服务器的用户名，也就是发送者的邮箱
smtp_pwd = "nwgbyyalowzxbceb"  # 授权码，和用户名user一起，用于登录smtp， 非邮箱密码
toEmails="wqliu_2008081171@163.com,wqliu163.163.com"

# 测试数据excel文件中，API表中列号映射,如“B”表示第二列
API_apiName = "B"
API_requestUrl = "C"
API_requestMothod = "D"
API_paramsType = "E"
API_apiTestCaseFileName = "F"
API_active = "G"

# 测试数据excel文件中，API的测试用例表中的列号数字映射
CASE_requestData = "A"
CASE_relyData = "B"
CASE_responseCode = "C"
CASE_responseData = "D"
CASE_checkPoint = "E"
CASE_active = "F"
CASE_status = "G"
CASE_failedReason = "H"
CASE_testTime="I"

#邮件开关：0-不发送，1-发送
switch=0