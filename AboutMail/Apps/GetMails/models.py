from django.db import models
from django.db import connection
import cx_Oracle as cx

# Create your models here.

# 邮件类
class EmailInfo:
    ID = ""
    MAILKEY = ""
    MAILDATE = ""
    MAILTIME = ""
    ENVFROM = ""
    ENVTO = ""
    HSUBJECT = ""
    ATTACH = ""
    MAIL_CONTENT_TEXT = ""
    MAIL_CONTENT_HTML = ""
    CRE_PERS = ""

    # 无参构造器
    def __init__(self):
        pass

# 数据库操作
class DBMS:
    # 将邮件信息插入数据库
    def InsertMail(self,EmailInfoInst):
        cursor = connection.cursor()

        clob_data1 = cursor.var(cx.CLOB)
        clob_data1.setvalue(0, "".join(EmailInfoInst.MAIL_CONTENT_HTML.split()).replace("'", "''"))
        clob_data2 = cursor.var(cx.CLOB)
        clob_data2.setvalue(0, "".join(EmailInfoInst.MAIL_CONTENT_TEXT.split()).replace("'", "''"))

        print("插入邮件表")
        strSql = "INSERT INTO DTAN_MAIL_INFO (ID,MAILKEY,MAILDATE,MAILTIME,ENVFROM,ENVTO,HSUBJECT,ATTACH,MAIL_CONTENT_TEXT,MAIL_CONTENT_HTML,REMARK,CRE_PERS" \
                 ") VALUES("
        strSql += "NVL('" + EmailInfoInst.ID.replace("'",
                                                     "''") + "','NULL'||SYS_GUID()),NVL('" + EmailInfoInst.MAILKEY.replace(
            "'",
            "''") + "','NULL'||SYS_GUID()),'" + EmailInfoInst.MAILDATE + "','" + EmailInfoInst.MAILTIME + "','" + EmailInfoInst.ENVFROM.replace(
            "'", "''") + "','" + EmailInfoInst.ENVTO.replace("'", "''") + "','" + EmailInfoInst.HSUBJECT.replace("'",
                                                                                                                 "''") + "','" + EmailInfoInst.ATTACH + "',:1,:2,'处理后','" + EmailInfoInst.CRE_PERS + "') "
        cursor.prepare(strSql)
        # perfect
        x = cursor.execute(None, {'1': " ".join(EmailInfoInst.MAIL_CONTENT_TEXT.split()).replace("'", "''"),
                             '2': (EmailInfoInst.MAIL_CONTENT_HTML).replace("'", "''")})
        connection.commit()

    # 将附件信息更新到数据库
    def UpdateMailAttach(self,strFileName,strID):
        cursor = connection.cursor()
        strSql = " UPDATE  DTAN_MAIL_INFO  SET ATTACH = '" + strFileName + "' WHERE ID = '" + strID + "' "
        x = cursor.execute(strSql)
        connection.commit()