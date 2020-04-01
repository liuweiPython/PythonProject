from django.db import models, connection


# Create your models here.

# 货Open信息
class CargoOpenInfo:
    RESULT_ID = ""
    MAIL_ID = ""
    CARGO_NAME = ""
    CARGO_VOLUME = ""
    UNIT = ""
    LOAD = ""
    DISCHARGE = ""
    LAYDAYS = ""
    CANCELING_DATE = ""
    LOADDISC_RATE = ""
    COMM = ""
    REMARKS = ""
    CRE_PERS = ""
    CRE_TIME = ""
    BEGIN = ""
    END = ""
    NO_CARGO = 0

    '''无参构造器'''
    def __init__(self):
        pass

# 数据库操作
class DBMS:
    # 获取邮件
    def GetMails(self):
        cursor = connection.cursor()
        strSql = " SELECT DMI.ID, DMI.HSUBJECT, DMI.MAIL_CONTENT_TEXT FROM DTAN_MAIL_INFO DMI WHERE 1 = 1 AND DMI.STYPE_C = '01'  AND DMI.MAIL_CONTENT_TEXT IS NOT NULL "
        x = cursor.execute(strSql)
        data = x.fetchall()
        return data
    # 更新邮件状态
    def UpdateFlag(cls, mail_id):
        cursor = connection.cursor()
        x = cursor.execute(" UPDATE DTAN_MAIL_INFO T SET T.FLAG_C= '01' WHERE  T.ID = '" + mail_id + "' ")
        connection.commit()
        pass
    # 插入分析结果
    def InsertCargoMailResult(cls, cargo, mail_id):
        cursor = connection.cursor()
        x = cursor.execute(" INSERT INTO CARGO_OPEN_RESULT_TEST(RESULT_ID,MAIL_ID,CARGO_NAME,CARGO_VOLUME,UNIT,LOAD,DISCHARGE,LAYDAYS,CANCELING_DATE,LOADDISC_RATE,COMM,CRE_PERS,CRE_TIME) VALUES(sys_guid(),'" + mail_id + "','" + cargo.CARGO_NAME + "','" + cargo.CARGO_VOLUME + "','" + cargo.UNIT + "','" + str(
                cargo.LOAD).rstrip(',') + "','" + str(cargo.DISCHARGE).rstrip(',') + "','" + str(cargo.LAYDAYS).rstrip(
                ',') + "','" + str(cargo.CANCELING_DATE).rstrip(
                ',') + "','" + cargo.LOADDISC_RATE + "','" + cargo.COMM + "','16274',SYSDATE) ")
        connection.commit()
        pass



