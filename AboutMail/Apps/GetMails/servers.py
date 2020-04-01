'''
业务逻辑写在这里
'''
from AboutMail.Apps.GetMails.models import EmailInfo, DBMS
from AboutMail.Apps.GetMails.utils import BaseCommon
import random
from email.parser import Parser
from email.utils import parseaddr
import poplib
import os
import sys
import traceback
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'

'''邮件处理类'''
class MailDeal:
    # 获取邮件信息
    def GetMails(self):
        # 读取文件计数
        fCount = open(r"C:\Count.txt", 'r+', encoding='utf-8')
        count = fCount.readline()
        if count == '':
            count = 1
        f = open(r"C:\MSGlog.txt", 'w+', encoding='utf-8')
        i = int(count)

        while i > 0:
            try:
                email = 'emailbak@trawind.com'
                password = 'AAaa00!!##'
                pop3_server = 'mail.trawind.com'
                server = poplib.POP3(pop3_server)
                # 可以打开或关闭调试信息:
                server.set_debuglevel(1)
                # 可选:打印POP3服务器的欢迎文字:
                print(server.getwelcome().decode('utf-8'))

                # 身份验证
                server.user(email)
                server.pass_(password)
                resp, lines, octets = server.retr(i)
                # lines存储了邮件原始文本的每一行
                # 可以获得整个邮件的原始文本.decode('utf-8')
                linesNew = []
                for i, val in enumerate(lines):
                    linesNew.append(bytes(BaseCommon.Rpl(str(val)[2:len(str(val)) - 1].replace(r'\xa0', r' ')), 'utf8'))

                msg_content = b'\r\n'.join(linesNew).decode('utf-8', 'ignore')
                # 稍后解析出邮件:
                msg = Parser().parsestr(msg_content)
                strID = self.GetMailSingle(msg, 0, f)

                if strID != "" and strID is not None:
                    print("删除邮件")
                    # 删除邮件
                    server.dele(i)
                    server.quit()
                i += 1
            except BaseException as e:
                i += 1
                fCount.seek(0)
                fCount.truncate()
                fCount.write(str(int(count) + 1))
                print("报错了")
                fErro = open(r"C:\Erro.txt", 'a+', encoding='utf-8')
                fErro.write("====================================")
                fErro.write(traceback.format_exc())

                # 退出程序
                sys.exit(1)
                pass
            continue

    # 单封获取邮件
    def GetMailSingle(self,msg, indent=0,f = None):
        EmailInfoInst = EmailInfo()
        # 如果是邮件进来,获取表头信息
        if indent == 0:
            for header in ['From', 'To', 'Subject', 'Date', 'Message-ID']:
                value = msg.get(header, '')
                if value:
                    # 主题
                    if header == 'Subject':
                        value = BaseCommon.decode_str(value)
                        EmailInfoInst.HSUBJECT = value
                    # 日期 'Sun, 9 Oct 2016 11:31:25 +0800'
                    elif header == 'Date':
                        value = BaseCommon.decode_str(value)
                        strDate = value[0:len(value) - 6]
                        format_time = BaseCommon.trans_format(strDate, '%a, %d %b %Y %H:%M:%S', '%Y-%m-%d %H:%M:%S')

                        datetime = format_time.split(' ')
                        EmailInfoInst.MAILTIME = datetime[1]
                        EmailInfoInst.MAILDATE = datetime[0]
                    # 发件人
                    elif header == 'From':
                        hdr, addr = parseaddr(value)
                        name = BaseCommon.decode_str(hdr)
                        value = u'%s <%s>' % (name, addr)
                        EmailInfoInst.ENVFROM = addr
                    # 收件人
                    else:
                        hdr, addr = parseaddr(value)
                        name = BaseCommon.decode_str(hdr)
                        value = u'%s <%s>' % (name, addr)
                        EmailInfoInst.ENVTO = addr
            # 内容处理
            self.GetContent(msg,EmailInfoInst)
            EmailInfoInst.CRE_PERS = '16274'
            EmailInfoInst.ID = "".join(random.sample('zyxwvutsrqponmlkjihgfedcba1230456789', 15))
            EmailInfoInst.MAILKEY = EmailInfoInst.ID

            str = EmailInfoInst.ID

            #邮件内容插入数据库
            DBMS.InsertMail(EmailInfoInst)

            # 获取附件
            self.GetFiles(msg, str)

            return str

    # 获取内容
    def GetContent(self,msg,EmailInfoInst):
        # 邮件是多个part
        if (msg.is_multipart()):
            parts = msg.get_payload()
            for n, part in enumerate(parts):
                self.GetContent(part,EmailInfoInst)
        else:
            content_type = msg.get_content_type()
            if content_type == 'text/html':
                content = msg.get_payload(decode=True)
                charset = BaseCommon.guess_charset(msg)
                if charset:
                    content = content.decode(charset, 'ignore')
                    # 替换标签后,TEXT
                    EmailInfoInst.MAIL_CONTENT_TEXT = BaseCommon.html_to_plain_text(content)
                    # HTML
                    EmailInfoInst.MAIL_CONTENT_HTML = content
            elif content_type == 'text/plain':
                if EmailInfoInst.MAIL_CONTENT_HTML == "":
                    content = msg.get_payload(decode=True)
                    charset = BaseCommon.guess_charset(msg)
                    if charset:
                        content = content.decode(charset, 'ignore')
                        # 替换标签后,TEXT
                        EmailInfoInst.MAIL_CONTENT_TEXT = BaseCommon.html_to_plain_text(content)
                        # HTML
                        EmailInfoInst.MAIL_CONTENT_HTML = content
        pass

    # 获取附件信息
    def GetFiles(self,msg,strID):
        strFileName = ""
        for part in msg.walk():
            filename = part.get_filename()
            if filename != None:  # 如果存在附件
                filename = BaseCommon.decode_str(filename)  # 获取的文件是乱码名称，通过一开始定义的函数解码
                data = part.get_payload(decode=True)  # 取出文件正文内容

                # 判断文件夹是否存在
                strPath = "E:\\ATTACH\\" + strID
                isExists = os.path.exists(strPath)
                if not isExists:
                    # 如果不存在则创建目录
                    # 生成文件夹
                    os.makedirs(strPath)

                # 此处可以自己定义文件保存位置
                path = strPath + "\\" + filename.replace(",", "").replace("\\", "")
                strFileName += filename.replace(",", "").replace("\\", "") + ","
                f = open(path, 'wb')
                f.write(data)
                f.close()

        # 保存附件信息到数据库
        DBMS.UpdateMailAttach(strFileName,strID)