from django.db import models
from django.db import connection
import cx_Oracle as cx

# Create your models here.

# 船OPEN信息
class VesselOpenInfo:
    RESULT_ID = ""
    MAIL_ID = ""
    VESSEL_ID = ""
    VESSEL_NAME = ""
    DWT = ""
    BUILT_YEAR = ""
    OPEN_AREA = ""
    OPEN_DATE_BEGIN = ""
    OPEN_DATE_END = ""
    VESSEL_NAME_SOURCE = ""
    DWT_SOURCE = ""
    DWCC_SOURCE = ""
    BUILT_YEAR_SOURCE = ""
    OPEN_AREA_SOURCE = ""
    OPEN_DATE_SOURCE = ""
    FLAG = ""
    CRE_PERS = ""
    CRE_TIME = ""
    '''无参构造器'''

    def __init__(self):
        pass

# 船信息
class VslInfo:
    VESSEL_NAME = ""
    FIRST_BEGIN = ""
    FIRST_END = ""
    SECOND_BEGIN = ""
    SECOND_END = ""
    THIRD_BEGIN = ""
    THIRD_END = ""
    '''无参构造器'''
    def __init__(self):
        pass

# 数据库操作
class DBMS:
    # 获取邮件
    def GetMails(self,strFlag):
        cursor = connection.cursor()
        # 主题邮件
        if strFlag == "S":
            strSql = " SELECT T.HSUBJECT,T.ID FROM DTAN_MAIL_INFO T WHERE T.STYPE = '01' AND T.HSUBJECT IS NOT NULL AND T.FLAG IS NULL "
        else:
            strSql = " SELECT  T.MAIL_CONTENT_TEXT,T.ID FROM DTAN_MAIL_INFO T WHERE T.STYPE = '01' AND  (T.FLAG IS NULL OR T.FLAG = '00' ) "
        x = cursor.execute(strSql)
        data = x.fetchall()
        return data

    # 更新邮件状态
    def UpdateMailFlag(self, mail_id, flag,strFlag):
        cursor = connection.cursor()
        if strFlag == "S":
            x = cursor.execute(" UPDATE DTAN_MAIL_INFO T SET T.FLAG= '"+flag+"' WHERE T.ID = '" + mail_id + "' ")
        else:
            x = cursor.execute(
                " UPDATE DTAN_MAIL_INFO T SET T.FLAG= '"+flag+"' WHERE (T.FLAG IS NULL OR T.FLAG = '00') AND T.ID = '" +
                mail_id + "'  ")
        connection.commit()
        pass

    # 更新船舶数量
    def UpdateVesselCount(cls, mail_id, sum):
        cursor = connection.cursor()
        x = cursor.execute(
            " UPDATE WZC_MAIL_ANALYZE_RESULT T SET T.SUM_SHIP= " + str(sum) + " WHERE T.MAIL_ID = '" + mail_id + "'  ")
        connection.commit()
        pass

    # 插入船OPEN分析结果
    def InsertVesselMailResult(cls, strMailID, openInfo):
        cursor = connection.cursor()
        x = cursor.execute(
            "INSERT INTO WZC_MAIL_ANALYZE_RESULT(RESULT_ID,MAIL_ID,VESSEL_NAME_SOURCE,OPEN_AREA_SOURCE,OPEN_DATE_SOURCE,DWT_SOURCE, BUILT_YEAR_SOURCE,CRE_PERS,CRE_TIME)VALUES(SYS_GUID(), '" + strMailID + "', '" + openInfo.VESSEL_NAME_SOURCE.rstrip(
                ',').replace("'", "''") + "', '" + openInfo.OPEN_AREA_SOURCE.rstrip(',').replace("'",
                                                                                                 "''") + "', '" + openInfo.OPEN_DATE_SOURCE.rstrip(
                ',').replace("'", "''") + "', '" + openInfo.DWT_SOURCE.rstrip(',').replace("'",
                                                                                           "''") + "','" + openInfo.BUILT_YEAR_SOURCE.rstrip(
                ',').replace("'", "''") + "','16274', SYSDATE) ")
        connection.commit()
        pass
