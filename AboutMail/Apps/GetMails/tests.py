
from django.test import TestCase

# Create your tests here.
import os
# 添加配置环境
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "AboutMail.AboutMail.settings")

# # 数据库测试
# from django.db import connection
#
# cursor = connection.cursor()
# cursor.execute("select * from hr_staff ")
# data = cursor.fetchall()
# i = 1

# # 定时任务测试
# import time
# import datetime
# from apscheduler.schedulers.blocking import BlockingScheduler
#
# def func():
#     now = datetime.datetime.now()
#     ts = now.strftime('%Y-%m-%d %H:%M:%S')
#     print('do func  time :', ts)
#
# def func2():
#     # 耗时2S
#     now = datetime.datetime.now()
#     ts = now.strftime('%Y-%m-%d %H:%M:%S')
#     print('do func2 time：', ts)
#     time.sleep(2)
#
# def dojob():
#     # 创建调度器：BlockingScheduler
#     scheduler = BlockingScheduler()
#     # 添加任务,时间间隔2S
#     scheduler.add_job(func, 'interval', seconds=2, id='test_job1')
#     # 添加任务,时间间隔5S
#     scheduler.add_job(func2, 'interval', seconds=3, id='test_job2')
#     scheduler.start()
#
# dojob()
