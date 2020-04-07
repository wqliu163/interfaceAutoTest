import os
import smtplib
import traceback
from email.mime.text import MIMEText
from email.mime.multipart import  MIMEMultipart
from email.mime.application import MIMEApplication
from config.config import *
from utils.log.excuteWriteLog import *

class sendEmail(object):
    smtp_port = smtp_port  # smtp服务器SSL端口号，默认是465
    smtp_host = smtp_host  # 发送邮件的smtp服务器（从QQ邮箱中取得）
    smtp_from_email = smtp_from_email  # 用于登录smtp服务器的用户名，也就是发送者的邮箱
    smtp_pwd = smtp_pwd  # 授权码，和用户名user一起，用于登录smtp， 非邮箱密码
    emailSuffix="gmail.com；@yahoo.com；@msn.com；@hotmail.com；@aol.com；@ask.com；@live.com；@qq.com；@0355.net；@163.com；@163.net；@263.net；@3721.net；@yeah"

    @classmethod
    def sendEmailNoCarry(cls,subject,bodys,toEmail=None):
        '''
        发送邮件
        param to_email_list: 收件人邮箱列表，格式"123@qq.com,123@163.com"
        param subject: 邮件主题，格式："邮件主题"
        param body:  邮件内容， 格式："邮件所说的内容"
        '''
        msg = MIMEMultipart()
        msg.attach(MIMEText(bodys, 'plain', 'utf-8'))
        msg["From"] = smtp_from_email  # 发件人
        msg["Subject"] = subject  # 邮件标题
        if toEmail==None:
            msg["To"] = toEmails  # 收件人列表,转换成string，用逗号隔开
        else:
            msg["To"] =toEmail
        try:
            SmtpSslClient = smtplib.SMTP_SSL(sendEmail.smtp_host, sendEmail.smtp_port)  # 实例化一个SMTP_SSL对象
            Loginer = SmtpSslClient.login(sendEmail.smtp_from_email, sendEmail.smtp_pwd)  # 登录smtp服务器
            #print("登录结果：Loginer=", Loginer)  # loginRes =  (235, b'Authentication successful')
            if Loginer[0] == 235:
                excutelog("info","Email login successful，login name is '%s'"%sendEmail.smtp_from_email)
                print("登录成功，code=", Loginer[0])
                if toEmail==None:
                    SmtpSslClient.sendmail(sendEmail.smtp_from_email, toEmails, msg.as_string())  # 发件人，收件人列表，邮件内容
                    excutelog("info", "send %s email successful" % toEmails)
                else:
                    SmtpSslClient.sendmail(sendEmail.smtp_from_email, toEmail, msg.as_string())  # 发件人，收件人列表，邮件内容
                    excutelog("info", "send %s email successful" % toEmail)
                excutelog("info","mail has been send successfully,message:%s"% msg.as_string())
                print("mail has been send successfully,message:", msg.as_string())
                SmtpSslClient.quit()  # 退出邮箱
            else:
                print("邮件登录失败，发送失败。code=", Loginer[0], "message=", msg.as_string())
        except:
            excutelog("error",traceback.print_exc())
            print("邮件发送失败，报错信息：", traceback.print_exc())

    @classmethod
    def sendEmailWithCarry(cls,subject,bodys,toEmail=None,filesPath=None):
        '''
        发送邮件
        param toEmail: 收件人邮箱列表，格式"123@qq.com,123@163.com"
        param subject: 邮件主题，格式："邮件主题"
        param body:  邮件内容， 格式："邮件所说的内容"
        param filesPath=None  发送的附件，默认不带附件，格式 r"E:\test.xlsx"
        '''
        msg = MIMEMultipart()
        msg.attach(MIMEText(bodys, 'plain', 'utf-8'))
        msg["From"] = sendEmail.smtp_from_email  # 发件人
        msg["Subject"] = subject  # 邮件标题
        if toEmail==None:
            msg["To"] = toEmails  # 收件人列表,转换成string，用逗号隔开
        else:
            msg["To"] =toEmail

        # 上传指定文件构造附件
        if filesPath!=None and os.path.isfile(filesPath) and os.path.exists(filesPath):
            filespart = MIMEApplication(open(filesPath, 'rb').read())
            file_name = filesPath.split("\\")[-1]  # 获取文件名,双\\转译\
            print("file_name=%s"%file_name)
            filespart.add_header("Content-Disposition", "attachment", filename=file_name)  # file_name是显示附件的名字，可随便自定义
            msg.attach(filespart)
        else:
            excutelog("info","加载的附件不存在!")
            print("加载的附件不存在")

        try:
            SmtpSslClient = smtplib.SMTP_SSL(sendEmail.smtp_host, sendEmail.smtp_port)  # 实例化一个SMTP_SSL对象
            Loginer = SmtpSslClient.login(sendEmail.smtp_from_email, sendEmail.smtp_pwd)  # 登录smtp服务器
            print("登录结果：Loginer=", Loginer)  # loginRes =  (235, b'Authentication successful')
            if Loginer[0] == 235:
                print("登录成功，code=", Loginer[0])
                excutelog("info", "Email login successful，login name is '%s'" % sendEmail.smtp_from_email)
                if toEmail==None:
                    SmtpSslClient.sendmail(sendEmail.smtp_from_email, toEmails, msg.as_string())  # 发件人，收件人列表，邮件内容
                    excutelog("info", "send %s email successful" % toEmails)
                else:
                    SmtpSslClient.sendmail(sendEmail.smtp_from_email, toEmail, msg.as_string())  # 发件人，收件人列表，邮件内容
                    excutelog("info","send %s email successful"%toEmail)
                print("mail has been send successfully,message:", msg.as_string())
                SmtpSslClient.quit()  # 退出邮箱
            else:
                print("邮件登录失败，发送失败。code=%s"%Loginer[0], "message=%s"%msg.as_string())
        except:
            excutelog("error", traceback.print_exc())
            print("邮件发送失败，报错信息：%s"%traceback.print_exc())

if __name__=="__main__":
    sendEmail.sendEmailNoCarry("邮件标题","邮件内容","wqliu163@163.com")
    #sendEmail.sendEmailNoCarry("邮件标题", "邮件内容")
    sendEmail.sendEmailWithCarry("带附件邮件","邮件内容",filesPath=r"C:\Users\liuwq\Desktop\新建文件夹\22号.txt")