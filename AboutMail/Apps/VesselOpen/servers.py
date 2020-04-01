'''
业务逻辑写在这里
'''
import jieba
jieba.set_dictionary("./dict.txt")
jieba.initialize()
jieba.load_userdict(r"C:\Dictionary.txt")
import jieba.posseg as pseg
import re
import os
from VesselOpen.models import DBMS, VslInfo, VesselOpenInfo

os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'

# 船OPEN邮件主题
class VesselOpenMailSubject:

    # 邮件主题分析
    def MailSubjectAnalyse(self):
        strFlag = "S"
        # 获取邮件信息
        data = DBMS.GetMails(strFlag)
        # 邮件计数
        i = 0
        # 遍历数据计数
        count = 0


        # 遍历邮件
        for mail in data:
            count += 1
            self.sum = 0

            DBMS.UpdateMailFlag( mail[1],'00',strFlag)

            mailInfo = mail[0].upper().replace("\r\n", " ");
            # 处理数字中的逗号分隔符
            pat = re.compile('\d{1,3}[,\d{3}]+')
            nms = re.findall(pat, mailInfo)
            if (len(nms) > 0):
                for nm in nms:
                    if ((",") in nm):
                        nmF = nm.replace(",", "")
                        mailInfo = mailInfo.replace(nm, nmF)
            # 开始分词
            words = pseg.cut(mailInfo)
            self.SingleFlag = True
            VslFlag = True
            i = 0
            # 一封邮件的关键词集合
            list = []
            # 位置，关键词，词性
            listC = []
            for word, flag in words:
                # 信息放入数组中
                if (flag != "x" and flag != "low" and flag != "n" and flag != "eng" and flag != "c"
                        and flag != "d" and flag != "nr" and flag != "v" and flag != "ns" or word == "K"):
                    i += 1
                    listC = []
                    listC.append(i)
                    listC.append(flag)
                    listC.append(word)
                    list.append(listC)
            # 邮件主题分析
            self.AnalyzerMailSubject(mail[1], list)

            # 更新船舶数量信息
            DBMS.UpdateVesselCount(mail[1],str(self.sum))

    # 邮件主题分析
    def AnalyzerMailSubject(self,strMailID, list):
        # 获取MV的序号
        NOmv = ""
        # open序号
        NOopen = ""
        # 第一次遍历
        for i, val in enumerate(list):
            if (val[1] == "vslkey"):
                # 判断MV
                NOmv += str(i) + ","
            elif (val[1] == "opkey"):
                # 判断OPEN
                NOopen += val[2] + ","
        # 上一个船名
        strLastVsl = ""
        # 创建数组存储船名对象
        vslList = []
        # 第二次遍历数据，将船名出现的开始和结束时间获取下来
        # 船名拼接字符串
        strVsl = ","
        for i, val in enumerate(list):
            # 船名第一次出現
            if (val[1] == "vsl" and not ("," + val[2] + "," in strVsl) and (
                    NOmv == "" or (NOmv != "" and (str(i - 1) in NOmv or len(NOopen) > 6)))):
                # 拼接船名
                strVsl += val[2] + ","
                # 創建船名對象
                vessel = VslInfo()
                # 賦值到第一次出現時間
                vessel.FIRST_BEGIN = val[0]
                vessel.VESSEL_NAME = val[2]
                # 添加船名到船名數組
                vslList.append(vessel)
                # 赋值END 如果上一艘船不为空
                if (strLastVsl != ""):
                    # 获取上一艘船在数组中的位置
                    numlistlast = strVsl.split(',')
                    numlast = numlistlast.index(strLastVsl) - 1
                    # 给上一艘船的结束位置赋值
                    if (vslList[numlast].FIRST_BEGIN != "" and vslList[numlast].FIRST_END == ""):
                        vslList[numlast].FIRST_END = val[0]
                # 上一个船名
                strLastVsl = val[2]

        # 将最后一个END赋值
        for i, val in enumerate(vslList):
            if (val.FIRST_BEGIN != "" and val.FIRST_END == ""):
                val.FIRST_END = list[len(list) - 1][0]

        # 第三次遍历获取船名区间内的时间关键字
        for j, vsl in enumerate(vslList):
            if (vsl.FIRST_BEGIN != "" and vsl.FIRST_END != ""):
                # 单条船处理
                self.AnalyzerVsl(vsl, list, strMailID)
        pass

    # 单条船分析处理
    def AnalyzerVsl(self, vsl, list, strMailID):
        # 船名序号
        NOvls = 0
        # 月份序号
        NOmonth = 0
        # 建造序号
        NOblt = ""
        # open序号
        NOopn = ""
        # flag 序号
        NOflag = 0
        # dwt序号
        NOdwt = 0
        # dwcc序号
        NOdwcc = 0
        # mt序号
        NOmt = 0
        # k 序号
        NOk = ""

        openInfo = VesselOpenInfo()
        openInfo.VESSEL_NAME_SOURCE = vsl.VESSEL_NAME
        # 遍历数据
        for i, val in enumerate(list):
            if (vsl.FIRST_BEGIN != "" and vsl.FIRST_END != ""):
                NOvls = vsl.FIRST_BEGIN
                if (i >= int(vsl.FIRST_BEGIN) and i < int(vsl.FIRST_END)):
                    if (val[1] == "monthdate" and i < int(vsl.FIRST_BEGIN) + 15 and NOmonth == 0):
                        # 判断月份
                        NOmonth = i
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
                    elif (val[1] == "btkey"):
                        # 判断建造序号
                        NOblt += str(i) + ","
                    elif (val[1] == "opkey"):
                        # open序号
                        NOopn += str(i) + ","
                    elif (val[1] == "flagkey"):
                        # flag序号
                        NOflag = i
                    elif (val[1] == "dwt"):
                        # dwt序号
                        NOdwt = i
                    elif (val[1] == "dwcc"):
                        # dwcc序号
                        NOdwcc = i
        # 获取mt和k
        for i, val in enumerate(list):
            if (val[1] == "mt"):
                # mt序号
                NOmt = i
            elif (val[2] == "K"):
                # k序号
                NOk += str(i) + ","

        dateflag = False
        # 获取主要信息
        for i, val in enumerate(list):
            if ((vsl.FIRST_BEGIN != "" and i >= int(vsl.FIRST_BEGIN) and i < int(vsl.FIRST_END))):
                # 获取OPEN区域
                if (val[1] == "port" or val[1] == "zone"):
                    openInfo.OPEN_AREA_SOURCE += val[2] + ","
                # 获取BUILT时间
                if (val[1] == "m" and len(val[2]) == 4 and ("199" in val[2] or "200" in val[2] or "201" in val[2]) and (
                        str(i + 1) in NOblt or str(i - 1) in NOblt)):
                    openInfo.BUILT_YEAR_SOURCE = val[2]
                # 获取OPEN时间
                if (vsl.FIRST_BEGIN != "" and i >= int(vsl.FIRST_BEGIN) and i < int(vsl.FIRST_END)):
                    if ((val[1] == "date" or val[1] == "daykey") and i > NOmonth - 3 and i < NOmonth + 3 and val[
                        1] != "time"):
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
                        dateflag = True
                    elif ((val[1] == "m" and (len(val[2]) < 3)) and i > NOmonth - 3 and i < NOmonth + 3 and val[
                        1] != "time"):
                        dateList = openInfo.OPEN_DATE_SOURCE.rstrip(',').split(",")
                        if (len(dateList) < 3):
                            openInfo.OPEN_DATE_SOURCE += val[2] + ","
                            dateflag = True
                    if (dateflag != True):
                        if ((val[1] == "date" or val[1] == "daykey") and i > NOmonth and i < NOmonth + 3 and val[
                            1] != "time"):
                            openInfo.OPEN_DATE_SOURCE += val[2] + ","
                        elif ((val[1] == "m" and len(val[2]) < 3) and i > NOmonth and i < NOmonth + 3 and val[
                            1] != "time"):
                            dateList = openInfo.OPEN_DATE_SOURCE.rstrip(',').split(",")
                            if (len(dateList) < 3):
                                openInfo.OPEN_DATE_SOURCE += val[2] + ","
                # 获取DWT
                if (val[1] == "m" and "." not in val[2] and (len(val[2]) <= 3) and str(
                        i + 1) in NOk and openInfo.DWT_SOURCE == ""):
                    openInfo.DWT_SOURCE = val[2] + "000"
                elif (val[1] == "m" and (
                        "." not in val[2] and (len(val[2]) >= 5 and len(val[2]) <= 7)) and openInfo.DWT_SOURCE == ""):
                    openInfo.DWT_SOURCE = val[2]
        if (
                openInfo.VESSEL_NAME_SOURCE != "" and openInfo.OPEN_DATE_SOURCE != "" and openInfo.OPEN_AREA_SOURCE != "" and len(
            openInfo.OPEN_DATE_SOURCE) < 50 and len(openInfo.OPEN_AREA_SOURCE) < 50):
            # 插入WZC_MAIL_ANALYZE_RESULT表中
            DBMS.InsertVesselMailResult(strMailID, openInfo)
            self.sum += 1

            '''更新邮件状态 mail_id'''
            DBMS.UpdateMailFlag(strMailID,"01","S")
        pass

# 船OPEN邮件内容
class VesselOpenMail:

    # 邮件内容分析
    def MailAnalyse(self):
        strFlag = ""
        # 邮件计数
        i = 0
        # 判断只一条船的标识
        self.SingleFlag = True
        # 获取邮件信息
        data = DBMS.GetMails(strFlag)
        '''遍历数据计数'''
        count = 0
        '''邮件船舶计数'''
        self.sum = 0
        '''遍历数据'''
        for mail in data:
            try:
                print()
                count += 1
                self.sum = 0
                print("邮件" + str(count) + ":" + mail[1] + "")

                mailInfo = mail[0].read().upper().replace("\r\n", " ");
                mailInfo = re.sub(' +', ' ', mailInfo)
                # 更新邮件状态
                DBMS.UpdateMailFlag( mail[1] , '03', strFlag)

                # 处理数字中的逗号分隔符
                pat = re.compile('\d{1,3}[,\d{3}]+')
                nms = re.findall(pat, mailInfo)
                if (len(nms) > 0):
                    for nm in nms:
                        if ((",") in nm):
                            nmF = nm.replace(",", "")
                            mailInfo = mailInfo.replace(nm, nmF)
                # 开始分词

                words = pseg.cut(mailInfo)
                self.SingleFlag = True
                VslFlag = True
                i = 0
                # 一封邮件的关键词集合
                list = []
                # 位置，关键词，词性
                listC = []
                for word, flag in words:
                    # 输出
                    if (flag != "x" and flag != "low" and flag != "n" and flag != "eng" and flag != "c"
                            and flag != "d" and flag != "nr" and flag != "v" and flag != "ns" or word == "K"):
                        i += 1
                        listC = []
                        listC.append(i)
                        listC.append(flag)
                        listC.append(word)
                        list.append(listC)
                # 多条船分析
                self.AnalyzerMailMul(mail[1], list)
                # 更新船舶数量信息
                DBMS.UpdateVesselCount(mail[1], str(self.sum))
            except BaseException:
                pass
            continue

    # 多条船分析
    def AnalyzerMailMul(self, strMailID, list):
        # 判断出所有船名及出现的位置
        # 船名拼接字符串
        strVsl = ","
        # 创建数组存储船名对象
        vslList = []

        # 上一个船名
        strLastVsl = ""
        # 第一次遍历数据，将船名出现的开始和结束时间获取下来
        for i, val in enumerate(list):
            # 船名第一次出現
            if (val[1] == "vsl" and not ("," + val[2] + "," in strVsl)):
                # 拼接船名
                strVsl += val[2] + ","
                # 創建船名對象
                vessel = VslInfo()
                # 賦值到第一次出現時間
                vessel.FIRST_BEGIN = val[0]
                vessel.VESSEL_NAME = val[2]
                # 添加船名到船名數組
                vslList.append(vessel)
                # 赋值END 如果上一艘船不为空
                if (strLastVsl != ""):
                    '''获取上一艘船在数组中的位置'''
                    numlistlast = strVsl.split(',')
                    numlast = numlistlast.index(strLastVsl) - 1
                    '''给上一艘船的结束位置赋值'''
                    if (vslList[numlast].FIRST_BEGIN != "" and vslList[numlast].FIRST_END == ""):
                        vslList[numlast].FIRST_END = val[0]
                    elif (vslList[numlast].SECOND_BEGIN != "" and vslList[numlast].SECOND_END == ""):
                        vslList[numlast].SECOND_END = val[0]
                    elif (vslList[numlast].THIRD_BEGIN != "" and vslList[numlast].THIRD_END == ""):
                        vslList[numlast].THIRD_END = val[0]
                '''上一个船名'''
                strLastVsl = val[2]
                '''船名不是第一次出现'''
            elif (val[1] == "vsl" and ("," + val[2] + "," in strVsl)):
                '''获取的是初始字符位置，考虑如何 结束 上一个时间'''
                numlist = strVsl.split(',')
                num = numlist.index(val[2]) - 1
                '''上一艘船不为空'''
                if (strLastVsl != ""):
                    '''获取上一艘船在数组中的位置'''
                    numlistlast = strVsl.split(',')
                    numlast = numlistlast.index(strLastVsl) - 1
                    if (vslList[numlast].FIRST_BEGIN != "" and vslList[numlast].FIRST_END == ""):
                        vslList[numlast].FIRST_END = val[0]
                    elif (vslList[numlast].SECOND_BEGIN != "" and vslList[numlast].SECOND_END == ""):
                        vslList[numlast].SECOND_END = val[0]
                    elif (vslList[numlast].THIRD_BEGIN != "" and vslList[numlast].THIRD_END == ""):
                        vslList[numlast].THIRD_END = val[0]
                if (vslList[num].SECOND_BEGIN == ""):
                    vslList[num].SECOND_BEGIN = val[0]
                elif ((vslList[num].SECOND_BEGIN != "" and vslList[num].THIRD_BEGIN == "")):
                    vslList[num].THIRD_BEGIN = val[0]
                '''上一个船名'''
                strLastVsl = val[2]
        '''将最后一个END赋值并输出数据'''
        for i, val in enumerate(vslList):
            if (val.FIRST_BEGIN != "" and val.FIRST_END == ""):
                val.FIRST_END = list[len(list) - 1][0]
            if (val.SECOND_BEGIN != "" and val.SECOND_END == ""):
                val.SECOND_END = list[len(list) - 1][0]
            if (val.THIRD_BEGIN != "" and val.THIRD_END == ""):
                val.THIRD_END = list[len(list) - 1][0]

        '''第二次遍历获取船名区间内的时间关键字  and vsl.VESSEL_NAME == 'LEDRA'''
        for j, vsl in enumerate(vslList):
            if (vsl.FIRST_BEGIN != "" and vsl.FIRST_END != ""):
                # 单条船处理
                self.AnalyzerVsl(vsl, list, strMailID)
        pass

    # 单条船处理
    def AnalyzerVsl(self, vsl, list, strMailID):
        # 船名序号
        NOvls = 0
        # 月份序号
        NOmonth = 0
        # 建造序号
        NOblt = ""
        # open序号
        NOopn = 0
        # flag 序号
        NOflag = 0
        # dwt序号
        NOdwt = 0
        # dwcc序号
        NOdwcc = 0
        # mt序号
        NOmt = 0
        # k 序号
        NOk = 0

        openInfo = VesselOpenInfo()
        openInfo.VESSEL_NAME_SOURCE = vsl.VESSEL_NAME

        # 遍历数据
        for i, val in enumerate(list):
            # Fisrt
            if (vsl.FIRST_BEGIN != "" and vsl.FIRST_END != ""):
                NOvls = vsl.FIRST_BEGIN
                if (i >= int(vsl.FIRST_BEGIN) and i <= int(vsl.FIRST_END)):
                    if (val[1] == "monthdate" and i < int(vsl.FIRST_BEGIN) + 15 and NOmonth == 0):
                        # 判断月份
                        NOmonth = i
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
                    elif (val[1] == "time"):
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
                    elif (val[1] == "btkey"):
                        # 判断建造序号
                        NOblt = str(i) + ","
                    elif (val[1] == "opkey"):
                        # open序号
                        NOopn = i
                    elif (val[1] == "flagkey"):
                        # flag序号
                        NOflag = i
                    elif (val[1] == "dwt"):
                        # dwt序号
                        NOdwt = i
                    elif (val[1] == "dwcc"):
                        # dwcc序号
                        NOdwcc = i
            # Second
            if (vsl.SECOND_BEGIN != "" and vsl.SECOND_END != ""):
                NOvls = vsl.SECOND_BEGIN
                if (i >= int(vsl.SECOND_BEGIN) and i <= int(vsl.SECOND_END)):
                    if (val[1] == "monthdate" and i < int(vsl.SECOND_BEGIN) + 15 and NOmonth == 0):
                        # 判断月份
                        NOmonth = i
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
                    elif (val[1] == "time"):
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
                    elif (val[1] == "btkey"):
                        # 判断建造序号
                        NOblt = str(i) + ","
                    elif (val[1] == "opkey"):
                        # open序号
                        NOopn = i
                    elif (val[1] == "flagkey"):
                        # flag序号
                        NOflag = i
                    elif (val[1] == "dwt"):
                        # dwt序号
                        NOdwt = i
                    elif (val[1] == "dwcc"):
                        # dwcc序号
                        NOdwcc = i
            # Third
            if (vsl.THIRD_BEGIN != "" and vsl.THIRD_END != ""):
                NOvls = vsl.THIRD_BEGIN
                if (i >= int(vsl.THIRD_BEGIN) and i <= int(vsl.THIRD_END)):
                    if (val[1] == "monthdate" and i < int(vsl.THIRD_BEGIN) + 15 and NOmonth == 0):
                        # 判断月份
                        NOmonth = i
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
                    elif (val[1] == "time"):
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
                    elif (val[1] == "btkey"):
                        # 判断建造序号
                        NOblt = str(i) + ","
                    elif (val[1] == "opkey"):
                        # open序号
                        NOopn = i
                    elif (val[1] == "flagkey"):
                        # flag序号
                        NOflag = i
                    elif (val[1] == "dwt"):
                        # dwt序号
                        NOdwt = i
                    elif (val[1] == "dwcc"):
                        # dwcc序号
                        NOdwcc = i
        # 获取mt和k
        for i, val in enumerate(list):
            if (val[1] == "mt" and ((i < NOdwt + 3 and i > NOdwt - 3) or (i < NOdwcc + 3 and i > NOdwcc - 3))):
                # mt序号
                NOmt = i
            elif (val[2] == "K" and ((i < NOdwt + 3 and i > NOdwt - 3) or (i < NOdwcc + 3 and i > NOdwcc - 3))):
                # k序号
                NOk = i

        # 判断时间（日期）
        dateflag = False;
        for i, val in enumerate(list):
            if (
                    (vsl.FIRST_BEGIN != "" and i >= int(vsl.FIRST_BEGIN) and i <= int(vsl.FIRST_END))
                    or (vsl.SECOND_BEGIN != "" and i >= int(vsl.SECOND_BEGIN) and i <= int(vsl.SECOND_END))
                    or (vsl.THIRD_BEGIN != "" and i >= int(vsl.THIRD_BEGIN) and i <= int(vsl.THIRD_END))
            ):
                if ((val[1] == "date" or val[1] == "daykey") and i > NOmonth - 3 and i < NOmonth + 3 and val[
                    1] != "time"):
                    openInfo.OPEN_DATE_SOURCE += val[2] + ","
                    dateflag = True
                elif ((val[1] == "m" and (len(val[2]) < 3)) and i > NOmonth - 3 and i < NOmonth + 3 and val[
                    1] != "time"):
                    dateList = openInfo.OPEN_DATE_SOURCE.split(",")
                    if (len(dateList) < 3):
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
                        dateflag = True
                if (dateflag != True):
                    if ((val[1] == "date" or val[1] == "daykey") and i > NOmonth and i < NOmonth + 3 and val[
                        1] != "time"):
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
                    elif ((val[1] == "m" and len(val[2]) < 3) and i > NOmonth and i < NOmonth + 3 and val[1] != "time"):
                        dateList = openInfo.OPEN_DATE_SOURCE.split(",")
                        if (len(dateList) < 3):
                            openInfo.OPEN_DATE_SOURCE += val[2] + ","
                # 判断区域 距离船名五个关键字并且不包括建造地点
                if (val[1] == "port" or val[1] == "zone"):
                    # 如果紧挨着open
                    if (NOopn != 0 and i == NOopn + 1):
                        openInfo.OPEN_AREA_SOURCE += val[2] + ","
                    else:
                        if (NOblt != ""):
                            bltList = NOblt.rstrip(",").split(",")
                            for j, blt in enumerate(bltList):
                                # 距离船名5个关键字以内并且不能挨着Flag
                                if (i < NOvls + 5 and ((int(blt) != 0 and (i > int(blt) + 2 or i < int(blt) - 2)) or (
                                        int(blt) == 0)) and (
                                        (NOflag != 0 and (i > NOflag + 1 or i < NOflag - 1)) or (NOflag == 0))):
                                    openInfo.OPEN_AREA_SOURCE += val[2] + ","
                        else:
                            if (i < NOvls + 12 and (
                                    (NOflag != 0 and (i > NOflag + 1 or i < NOflag - 1)) or (NOflag == 0))):
                                openInfo.OPEN_AREA_SOURCE += val[2] + ","
                # 获取DWT
                if (NOdwt != 0 and val[1] == "m" and (
                        (len(val[2]) >= 2 and '.' not in val[2]) or ((len(val[2]) >= 6) and ('.' in val[2]))) and (
                        i > NOdwt - 3 and i < NOdwt + 3)):
                    if (NOblt != ""):
                        bltList = NOblt.rstrip(",").split(",")
                        for j, blt in enumerate(bltList):
                            if ((NOmt != 0 and i == NOmt - 1 and i != int(blt) - 1 and i != int(blt) + 1) or (
                                    NOmt != 0 and i == NOmt - 1 and ((i == int(blt) - 1 or i == int(blt) + 1) and (
                                    len(val[2]) != 4 or (
                                    len(val[2]) == 4 and "200" not in val[2] and "201" not in val[2] and "199" not in
                                    val[2]))))):
                                openInfo.DWT_SOURCE = val[2]
                            else:
                                if ((openInfo.DWT_SOURCE == "" and i != int(blt) - 1 and i != int(blt) + 1) or (
                                        openInfo.DWT_SOURCE == "" and ((i == int(blt) - 1 or i == int(blt) + 1) and (
                                        len(val[2]) != 4 or (
                                        len(val[2]) == 4 and "200" not in val[2] and "201" not in val[
                                    2] and "199" not in val[2]))))):
                                    openInfo.DWT_SOURCE = val[2]
                            if ((NOk != 0 and i == NOk - 1 and i != int(blt) - 1 and i != int(blt) + 1) or (
                                    NOk != 0 and i == NOk - 1 and ((i == int(blt) - 1 or i == int(blt) + 1) and (
                                    len(val[2]) != 4 or (
                                    len(val[2]) == 4 and "200" not in val[2] and "201" not in val[2] and "199" not in
                                    val[2]))))):
                                openInfo.DWT_SOURCE = val[2] + "000"
                            else:
                                if ((openInfo.DWT_SOURCE == "" and i != int(blt) - 1 and i != int(blt) + 1) or (
                                        openInfo.DWT_SOURCE == "" and ((i == int(blt) - 1 or i == int(blt) + 1) and (
                                        len(val[2]) != 4 or (
                                        len(val[2]) == 4 and "200" not in val[2] and "201" not in val[
                                    2] and "199" not in val[2]))))):
                                    openInfo.DWT_SOURCE = val[2]
                # 获取DWCC
                if (NOdwcc != 0 and val[1] == "m" and (
                        (len(val[2]) >= 2 and '.' not in val[2]) or ((len(val[2]) >= 6) and ('.' in val[2]))) and (
                        i > NOdwcc - 3 and i < NOdwcc + 3)):
                    if (NOblt != ""):
                        bltList = NOblt.rstrip(",").split(",")
                        for j, blt in enumerate(bltList):
                            if ((NOmt != 0 and i == NOmt - 1 and i != int(blt) - 1 and i != int(blt) + 1) or (
                                    NOmt != 0 and i == NOmt - 1 and (openInfo.DWT_SOURCE == "" and (
                                    (i == int(blt) - 1 or i == int(blt) + 1) and (len(val[2]) != 4 or (
                                    len(val[2]) == 4 and "201" not in val[2] and "200" not in val[2] and "199" not in
                                    val[2])))))):
                                openInfo.DWCC_SOURCE = val[2]
                            else:
                                if ((openInfo.DWCC_SOURCE == "" and i != int(blt) - 1 and i != int(blt) + 1) or (
                                        openInfo.DWCC_SOURCE == "" and (openInfo.DWT_SOURCE == "" and (
                                        (i == int(blt) - 1 or i == int(blt) + 1) and (len(val[2]) != 4 or (
                                        len(val[2]) == 4 and "201" not in val[2] and "200" not in val[
                                    2] and "199" not in val[2])))))):
                                    openInfo.DWCC_SOURCE = val[2]
                            if ((NOk != 0 and i == NOk - 1 and i != int(blt) - 1 and i != int(blt) + 1) or (
                                    NOk != 0 and i == NOk - 1 and (openInfo.DWT_SOURCE == "" and (
                                    (i == int(blt) - 1 or i == int(blt) + 1) and (len(val[2]) != 4 or (
                                    len(val[2]) == 4 and "201" not in val[2] and "200" not in val[2] and "199" not in
                                    val[2])))))):
                                openInfo.DWCC_SOURCE = val[2] + "000"
                            else:
                                if ((openInfo.DWCC_SOURCE == "" and i != int(blt) - 1 and i != int(blt) + 1) or (
                                        openInfo.DWCC_SOURCE == "" and (openInfo.DWT_SOURCE == "" and (
                                        (i == int(blt) - 1 or i == int(blt) + 1) and (len(val[2]) != 4 or (
                                        len(val[2]) == 4 and "201" not in val[2] and "200" not in val[
                                    2] and "199" not in val[2])))))):
                                    openInfo.DWCC_SOURCE = val[2]
                # 获取BUILT时间
                if (val[1] == "m" and len(val[2]) == 4 and ("199" in val[2] or "200" in val[2] or "201" in val[2])):
                    if (NOblt != ""):
                        bltList = NOblt.rstrip(",").split(",")
                        for j, blt in enumerate(bltList):
                            if (i == int(blt) + 1 or i == int(blt) - 1 or i == int(blt) + 2 or i == int(blt) - 2):
                                openInfo.BUILT_YEAR_SOURCE = val[2]
        # 船名在后Flag
        global VslFlag
        if (
                openInfo.VESSEL_NAME_SOURCE != "" and openInfo.OPEN_DATE_SOURCE != "" and openInfo.OPEN_AREA_SOURCE != "" and len(
                openInfo.OPEN_DATE_SOURCE) < 50 and len(openInfo.OPEN_AREA_SOURCE) < 100):

            # 插入WZC_MAIL_ANALYZE_RESULT表中
            DBMS.InsertVesselMailResult(strMailID,openInfo)

            self.sum += 1

            # 更新邮件状态
            DBMS.UpdateMailFlag(strMailID,"02","")

            VslFlag = False
        if (VslFlag and self.SingleFlag):
            # 船名在後的郵件分析
            self.AnalyzerVslBHD(list, strMailID, vsl)

        pass

    # 船名在後的郵件分析
    def AnalyzerVslBHD(self, list, strMailID, vsl):
        # 船名序号
        NOvls = vsl.FIRST_BEGIN
        # 月份序号
        NOmonth = 0
        # 建造序号
        NOblt = ""
        # open序号
        NOopn = 0
        # flag 序号
        NOflag = 0
        # dwt序号
        NOdwt = 0
        # dwcc序号
        NOdwcc = 0
        # k序号
        NOk = 0
        # mt序号
        NOmt = 0

        openInfo = VesselOpenInfo()

        # 第一次遍历
        for i, val in enumerate(list):
            # 判斷船名,不用考慮多次出現
            if (val[1] == "vsl" and val[0] == NOvls):
                openInfo.VESSEL_NAME_SOURCE = val[2]

            if (val[1] == "monthdate" and i < int(NOvls) and NOmonth == 0):
                # 判断月份
                NOmonth = i
                openInfo.OPEN_DATE_SOURCE += val[2] + ","
            elif (val[1] == "time"):
                openInfo.OPEN_DATE_SOURCE += val[2] + ","
            elif (val[1] == "btkey"):
                # 判断建造序号
                NOblt = str(i) + ","
            elif (val[1] == "opkey"):
                # open序号
                NOopn = i
            elif (val[1] == "flagkey"):
                # flag序号
                NOflag = i
            elif (val[1] == "dwt"):
                # dwt 序号
                NOdwt = i
            elif (val[1] == "dwcc"):
                # dwcc 序号
                NOdwcc = i
        # 获取mt和k
        for i, val in enumerate(list):
            if (val[1] == "mt" and ((i < NOdwt + 3 and i > NOdwt - 3) or (i < NOdwcc + 3 and i > NOdwcc - 3))):
                # mt序号
                NOmt = i
            elif (val[2] == "K" and ((i < NOdwt + 3 and i > NOdwt - 3) or (i < NOdwcc + 3 and i > NOdwcc - 3))):
                # k序号
                NOk = i

        # 第二次遍历
        dateflag = False
        for i, val in enumerate(list):
            if ((val[1] == "date" or val[1] == "daykey") and i > NOmonth - 3 and i < NOmonth + 3 and val[1] != "time"):
                openInfo.OPEN_DATE_SOURCE += val[2] + ","
                dateflag = True
            elif ((val[1] == "m" and len(val[2]) < 3) and i > NOmonth - 3 and i < NOmonth + 3 and val[1] != "time"):
                dateList = openInfo.OPEN_DATE_SOURCE.split(",")
                if (len(dateList) < 3):
                    openInfo.OPEN_DATE_SOURCE += val[2] + ","
                    dateflag = True
            if (dateflag != True):
                if ((val[1] == "date" or val[1] == "daykey") and i > NOmonth and i < NOmonth + 3 and val[1] != "time"):
                    openInfo.OPEN_DATE_SOURCE += val[2] + ","
                elif ((val[1] == "m" and len(val[2]) < 3) and i > NOmonth and i < NOmonth + 3 and val[1] != "time"):
                    dateList = openInfo.OPEN_DATE_SOURCE.split(",")
                    if (len(dateList) < 3):
                        openInfo.OPEN_DATE_SOURCE += val[2] + ","
            # 判断区域  距离船名五个关键字并且不包括建造地点
            if (val[1] == "port" or val[1] == "zone"):
                # 如果紧挨着open
                if (NOopn != 0 and i == NOopn + 1):
                    openInfo.OPEN_AREA_SOURCE += val[2] + ","
                else:
                    if (NOblt != ""):
                        bltList = NOblt.rstrip(",").split(",")
                        for j, blt in enumerate(bltList):
                            # 距离船名5个关键字以内并且不能挨着Flag
                            if (i < NOvls and (
                                    (int(blt) != 0 and (i > int(blt) + 2 or i < int(blt) - 2)) or (int(blt) == 0)) and (
                                    (NOflag != 0 and (i > NOflag + 1 or i < NOflag - 1)) or (NOflag == 0))):
                                openInfo.OPEN_AREA_SOURCE += val[2] + ","
                    else:
                        if (i < NOvls and (
                                (NOflag != 0 and (i > NOflag + 1 or i < NOflag - 1)) or (NOflag == 0))):
                            openInfo.OPEN_AREA_SOURCE += val[2] + ","
            # 获取DWT
            if (NOdwt != 0 and val[1] == "m" and (
                    (len(val[2]) >= 2 and '.' not in val[2]) or ((len(val[2]) >= 6) and ('.' in val[2]))) and (
                    i > NOdwt - 3 and i < NOdwt + 3)):
                if (NOblt != ""):
                    bltList = NOblt.rstrip(",").split(",")
                    for j, blt in enumerate(bltList):
                        if ((NOmt != 0 and i == NOmt - 1 and i != int(blt) - 1 and i != int(blt) + 1) or (
                                NOmt != 0 and i == NOmt - 1 and ((i == int(blt) - 1 or i == int(blt) + 1) and (
                                len(val[2]) != 4 or (
                                len(val[2]) == 4 and "201" not in val[2] and "200" not in val[2] and "199" not in val[
                            2]))))):
                            openInfo.DWT_SOURCE = val[2]
                        else:
                            if ((openInfo.DWT_SOURCE == "" and i != int(blt) - 1 and i != int(blt) + 1) or (
                                    openInfo.DWT_SOURCE == "" and ((i == int(blt) - 1 or i == int(blt) + 1) and (
                                    len(val[2]) != 4 or (
                                    len(val[2]) == 4 and "201" not in val[2] and "200" not in val[2] and "199" not in
                                    val[2]))))):
                                openInfo.DWT_SOURCE = val[2]
                        if ((NOk != 0 and i == NOk - 1 and i != int(blt) - 1 and i != int(blt) + 1) or (
                                NOk != 0 and i == NOk - 1 and ((i == int(blt) - 1 or i == int(blt) + 1) and (
                                len(val[2]) != 4 or (
                                len(val[2]) == 4 and "201" not in val[2] and "200" not in val[2] and "199" not in val[
                            2]))))):
                            openInfo.DWT_SOURCE = val[2] + "000"
                        else:
                            if ((openInfo.DWT_SOURCE == "" and i != int(blt) - 1 and i != int(blt) + 1) or (
                                    openInfo.DWT_SOURCE == "" and ((i == int(blt) - 1 or i == int(blt) + 1) and (
                                    len(val[2]) != 4 or (
                                    len(val[2]) == 4 and "201" not in val[2] and "200" not in val[2] and "199" not in
                                    val[2]))))):
                                openInfo.DWT_SOURCE = val[2]
                else:
                    if (NOmt != 0 and i == NOmt - 1):
                        openInfo.DWT_SOURCE = val[2]
                    else:
                        if (openInfo.DWT_SOURCE == ""):
                            openInfo.DWT_SOURCE = val[2]
                    if (NOk != 0 and i == NOk - 1):
                        openInfo.DWT_SOURCE = val[2] + "000"
                    else:
                        if (openInfo.DWT_SOURCE == ""):
                            openInfo.DWT_SOURCE = val[2]
            # 获取DWCC
            if (NOdwcc != 0 and val[1] == "m" and (
                    (len(val[2]) >= 2 and '.' not in val[2]) or ((len(val[2]) >= 6) and ('.' in val[2]))) and (
                    i > NOdwcc - 3 and i < NOdwcc + 3)):
                if (NOblt != ""):
                    bltList = NOblt.rstrip(",").split(",")
                    for j, blt in enumerate(bltList):
                        if ((NOmt != 0 and i == NOmt - 1 and i != int(blt) - 1 and i != int(blt) + 1) or (
                                NOmt != 0 and i == NOmt - 1 and (openInfo.DWT_SOURCE == "" and (
                                (i == int(blt) - 1 or i == int(blt) + 1) and (
                                len(val[2]) != 4 or (
                                len(val[2]) == 4 and "201" not in val[2] and "200" not in val[2] and "199" not in val[
                            2])))))):
                            openInfo.DWCC_SOURCE = val[2]
                        else:
                            if ((openInfo.DWCC_SOURCE == "" and i != int(blt) - 1 and i != int(blt) + 1) or (
                                    openInfo.DWCC_SOURCE == "" and (openInfo.DWT_SOURCE == "" and (
                                    (i == int(blt) - 1 or i == int(blt) + 1) and (len(val[2]) != 4 or (
                                    len(val[2]) == 4 and "201" not in val[2] and "200" not in val[2] and "199" not in
                                    val[2])))))):
                                openInfo.DWCC_SOURCE = val[2]
                        if ((NOk != 0 and i == NOk - 1 and i != int(blt) - 1 and i != int(blt) + 1) or (
                                NOk != 0 and i == NOk - 1 and (openInfo.DWT_SOURCE == "" and (
                                (i == int(blt) - 1 or i == int(blt) + 1) and (
                                len(val[2]) != 4 or (
                                len(val[2]) == 4 and "201" not in val[2] and "200" not in val[2] and "199" not in val[
                            2])))))):
                            openInfo.DWCC_SOURCE = val[2] + "000"
                        else:
                            if ((openInfo.DWCC_SOURCE == "" and i != int(blt) - 1 and i != int(blt) + 1) or (
                                    openInfo.DWCC_SOURCE == "" and (openInfo.DWT_SOURCE == "" and (
                                    (i == int(blt) - 1 or i == int(blt) + 1) and (len(val[2]) != 4 or (
                                    len(val[2]) == 4 and "201" not in val[2] and "200" not in val[2] and "199" not in
                                    val[2])))))):
                                openInfo.DWCC_SOURCE = val[2]
                else:
                    if (NOmt != 0 and i == NOmt - 1):
                        openInfo.DWCC_SOURCE = val[2]
                    else:
                        if (openInfo.DWCC_SOURCE == ""):
                            openInfo.DWCC_SOURCE = val[2]
                    if (NOk != 0 and i == NOk - 1):
                        openInfo.DWCC_SOURCE = val[2] + "000"
                    else:
                        if (openInfo.DWCC_SOURCE == ""):
                            openInfo.DWCC_SOURCE = val[2]
            # 获取BUILT时间
            if (val[1] == "m" and len(val[2]) == 4 and ("199" in val[2] or "200" in val[2] or "201" in val[2])):
                if (NOblt != ""):
                    bltList = NOblt.rstrip(",").split(",")
                    for j, blt in enumerate(bltList):
                        if (i == int(blt) + 1 or i == int(blt) - 1 or i == int(blt) + 2 or i == int(blt) - 2):
                            openInfo.BUILT_YEAR_SOURCE = val[2]

        if (
                openInfo.VESSEL_NAME_SOURCE != "" and openInfo.OPEN_DATE_SOURCE != "" and openInfo.OPEN_AREA_SOURCE != "" and len(
                openInfo.OPEN_DATE_SOURCE) < 18 and len(openInfo.OPEN_AREA_SOURCE) < 30):
            # 插入WZC_MAIL_ANALYZE_RESULT表中
            DBMS.InsertVesselMailResult(strMailID, openInfo)

            self.sum += 1

            # 更新邮件状态
            DBMS.UpdateMailFlag(strMailID, "02", "")
        self.SingleFlag = False
        pass