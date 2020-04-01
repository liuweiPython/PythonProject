import re
import jieba.posseg as pseg
import jieba
import traceback
import os

from CargoOpen.models import DBMS, CargoOpenInfo

jieba.initialize()
jieba.load_userdict(r"C:\CargoDictionary.txt")

os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'

# 货OPEN类
class CargoOpenMail:
    def MailAnalyse(self):
        # 获取邮件信息
        data = DBMS.GetMails()

        # 遍历数据
        for mail in data:
            try:
                mailInfo = mail[2].read().upper().replace("\r\n", " ")
                mailInfo = re.sub(' +', ' ', mailInfo)
                # 更新邮件状态
                DBMS.UpdateFlag(mail[0])
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
                # 排序
                i = 0
                # 一封邮件的关键词集合
                list = []
                # 位置，关键词，词性
                listC = []
                for word, flag in words:
                    # if 1==1:
                    if (flag != 'x' and flag != "low" and flag != "n" and flag != "eng" and flag != "c"
                            and flag != "d" and flag != "nr" and flag != "v" and flag != "ns" or word == "K"):
                        i += 1
                        listC = []
                        listC.append(i)
                        listC.append(flag)
                        listC.append(word)
                        list.append(listC)
                    elif (flag == 'x'):
                        if ((
                                word == "/" or word == "-" or word == "- " or word == ' – ' or word == '–' or word == '– ' or word == ' –' or word == " - " or word == " -" or word == " /" or word == " / " or word == "/ ")):
                            i += 1
                            listC = []
                            listC.append(i)
                            listC.append("split")
                            listC.append(word)
                            list.append(listC)
                        elif (word == "MARC"):
                            i += 1
                            listC = []
                            listC.append(i)
                            listC.append("monthdate")
                            listC.append("MAR")
                            list.append(listC)
                # 分析货邮件
                self.AnalyzerCargoMail(mail[0], list)

            except BaseException as e:
                print(traceback.format_exc())
                pass
            continue

    # 分析货邮件
    def AnalyzerCargoMail(self, mail_id, list):
        # 货对象的集合
        cargoList = []
        # 第一次遍历,确定货名和起始位置
        for i, val in enumerate(list):
            if val[1] == "cargo":
                cargo = CargoOpenInfo()
                if val[0] - 8 > 0:
                    cargo.BEGIN = val[0] - 8
                else:
                    cargo.BEGIN = 0
                cargo.CARGO_NAME = val[2]
                cargo.NO_CARGO = i
                cargoList.append(cargo)
                if len(cargoList) >= 2:
                    cargoList[len(cargoList) - 2].END = val[0] - 7
        if len(cargoList) > 0:
            # 最后一个END
            if int(list[len(list) - 1][0]) - int(cargoList[len(cargoList) - 1].BEGIN) < 60:
                cargoList[len(cargoList) - 1].END = list[len(list) - 1][0]
            else:
                cargoList[len(cargoList) - 1].END = int(cargoList[len(cargoList) - 1].BEGIN) + 60
        for j, cargo in enumerate(cargoList):
            if (cargo.BEGIN != "" and cargo.END != ""):
                # 单条船处理
                self.AnalyzerCargo(cargo, list, mail_id)
        pass

    # 单条船处理
    def AnalyzerCargo(self, cargo, list, mail_id):
        noCargo = cargo.NO_CARGO
        # 单位编号
        noUnit = 0
        # laycan
        noLaycan = 0
        # percent
        noPercent = 0
        # 月份
        noMonth = 0
        # 佣金
        noComm = 0
        # 日期分隔符
        noDateSplit = 0
        # 分隔符拼接
        strSplit = ""
        # 佣金百分号
        strCommonPercet = ""
        # first
        noFirst = 0
        # second
        noSecond = 0
        # port
        strPort = ""
        # 地点分隔符
        noPortSplit = 0
        # 时间类型
        timeFlag = False
        toFlag = False
        # 第一次遍历
        for i, val in enumerate(list):
            if i >= cargo.BEGIN and i <= cargo.END:

                # 加上在月份之后这个条件加上单位为0这个条件
                if val[1] == "unit" and (noUnit == 0 or (noUnit != 0 and i < noCargo)) and i <= cargo.BEGIN + 20 and (
                        noMonth == 0 or (noMonth > 0 and i < noMonth)):
                    noUnit = i
                    cargo.UNIT = val[2]
                elif val[1] == "laycan":
                    noLaycan = i
                elif val[1] == "monthdate" and (
                        noLaycan == 0 or (noLaycan != 0 and i > noLaycan)) and noMonth == 0 and i > noCargo:
                    if (not timeFlag):
                        noMonth = i
                        cargo.LAYDAYS += val[2]
                        cargo.CANCELING_DATE += val[2]
                elif val[1] == "time" and (
                        noLaycan == 0 or (noLaycan != 0 and i > noLaycan)) and noMonth == 0 and i > noCargo:
                    timeFlag = True
                    cargo.LAYDAYS += val[2]
                    cargo.CANCELING_DATE += val[2]
                elif val[1] == "comm":
                    noComm = i
                elif val[1] == "portflag":
                    if noFirst == 0:
                        noFirst = i
                    else:
                        noSecond = i
                elif val[1] == "zone" or val[1] == "port":
                    strPort += str(i) + ","

        # 如果都等于0
        if noFirst == 0 or noSecond == 0:
            noFirst = 0
            noSecond = 0
            for i, val in enumerate(list):
                if i >= cargo.BEGIN and i <= cargo.END:
                    if val[1] == "lflag" and noFirst == 0 and noSecond == 0:
                        if noFirst == 0:
                            noFirst = i
                        if str(val[2]) == "FM":
                            toFlag = True
                    elif val[1] == "dflag" and noSecond == 0 and noFirst != 0:
                        if ((toFlag and str(val[2]) == "TO") or (not toFlag and str(val[2]) != "TO")):
                            noSecond = i
                    elif (str(val[2]) == "/" or str(val[2]) == " /" or str(val[2]) == "/ " or str(
                            val[2]) == " / " or str(val[2]) == ' – ' or str(val[2]) == " – " or str(
                            val[2]) == "– " or str(val[2]) == " –" or str(val[2]) == "–") and (
                            str(i - 1) in strPort and (
                            str(i + 1) in strPort or str(i + 2) in strPort or str(i + 3) in strPort)):
                        noPortSplit = i
        # 第二次遍历
        for i, val in enumerate(list):
            if i >= cargo.BEGIN and i <= cargo.END:
                if noUnit != 0:
                    if i < noUnit and i > noUnit - 3 and val[1] == "m":
                        cargo.CARGO_VOLUME = val[2]
                else:
                    if val[1] == "m" and (i > noCargo - 5 and i < noCargo):
                        cargo.CARGO_VOLUME = val[2]
                # 装卸港判断
                if (val[1] == "zone" or val[1] == "port"):
                    if noFirst > 0 and noSecond > 0 and (int(noSecond) - int(noFirst)) <= 13:
                        if (i > noFirst and i < noSecond):
                            cargo.LOAD += str(val[2]) + ","
                        elif (i > noSecond and i < noSecond + 13):
                            cargo.DISCHARGE += str(val[2]) + ","
                    elif noPortSplit > 0:
                        if (i > noPortSplit - 10 and i < noPortSplit):
                            cargo.LOAD += str(val[2]) + ","
                        elif (i > noPortSplit and i < noPortSplit + 13):
                            cargo.DISCHARGE += str(val[2]) + ","
                elif ((val[1] == "date" or (val[1] == "m" and len(str(val[2])) < 3) or val[1] == "split") and (
                        (noLaycan == 0 and i >= noMonth - 6 and i < noMonth) or (
                        noLaycan != 0 and i > noLaycan and i <= noLaycan + 7))):
                    if str(val[2]) != "." and "2020" not in str(cargo.LAYDAYS) and "2019" not in str(
                            cargo.LAYDAYS) and not timeFlag:
                        cargo.LAYDAYS += str(val[2]) + " "
                        cargo.CANCELING_DATE += str(val[2]) + " "

                elif (val[1] == 'm' and i < noComm + 3 and i > noComm - 3 and str(val[2]) != "2020" and str(
                        val[2]) != "2019"):
                    cargo.COMM = val[2]
                elif (val[2] == "%" and i < noComm + 3 and i > noComm - 3):
                    strCommonPercet = "%"
        # 插入结果表
        if (cargo.CARGO_VOLUME != "" and cargo.LOAD != "" and str(cargo.CARGO_VOLUME) != "." and len(
                str(cargo.CARGO_VOLUME)) < 10 and len(str(
                cargo.LOAD)) < 60 and cargo.DISCHARGE != "" and cargo.CARGO_NAME != "" and cargo.CANCELING_DATE != "" and cargo.LAYDAYS != ""):
            DBMS.InsertCargoMailResult(cargo, mail_id)
        pass