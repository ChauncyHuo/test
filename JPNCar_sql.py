# -*- coding: utf-8 -*-
import time
import pymssql
import socket
import sys
import win32com.client as win32
import pythoncom
import importlib
from JPNCar_data import Dict_L2Group
#自定义执行函数
class WriteClass():
    def __init__(self):
        self.__conn=pymssql.connect(host=r"xmntsdb06.apac.dell.com\drt",user=r"AutoAgent",password=r"Auto__Agent",charset=r"utf8",database='DPSAUTO')
        self.__cur=self.__conn.cursor()
        self.machine_id = socket.getfqdn(socket.gethostname()).split(".")[0]
        
    def GetDB(self,sql):
        self.__cur.execute(sql)
        rows=self.__cur.fetchall()
        return rows
    
    def UpdateDB(self,sql):
        self.__cur.execute(sql,)
        self.__conn.commit()
    
    def UpdateAddGetDB(self,sql):
        self.__cur.execute(sql,)
        rows=self.__cur.fetchall()
        self.__conn.commit()
        return rows
    
    def close(self):
        self.__conn.close()
        
    def WriteData(self,table,id_cache):
        if table['status'].strip() not in ['approve','Pending']: 
            Sendmail(table)
        else:
            pass
        recordtime=format(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
        sql_update="""update [DPSAUTO].[dbo].[JPN_Autoapprove_Record] set
        [break_fix_id]=N'{}',[status]=N'{}',[reason]=N'{}',[rework number]=N'{}',[process time]=N'{}',[DPS#]=N'{}'
        where [break_fix_id]=N'{}'""".format(table['break_fix_id'],table['status'],
        table['reason'],table['rework number'],recordtime,table['DPS#'],id_cache)
        self.__cur.execute(sql_update,)
        self.__conn.commit()

class Sql_Class():
    def __init__(self,host_num):
        if host_num == 6:
            host = 'XMNTSDB06.apac.dell.com\drt'
            user = 'AutoAgent'
            password = 'Auto__Agent'
        elif host_num == 2:
            host = 'xmntsdb02.apac.dell.com'
            user = 'TCD_User'
            password = 'TCD_123'
        self.__conn=pymssql.connect(host,user,password,charset=r"utf8")
        self.__cur=self.__conn.cursor()
        self.machine_id = socket.getfqdn(socket.gethostname()).split(".")[0]
        
    def GetDB(self,sql):
        self.__cur.execute(sql)
        rows=self.__cur.fetchall()
        return rows
    
    def UpdateDB(self,sql):
        self.__cur.execute(sql,)
        self.__conn.commit()
    
    def UpdateAddGetDB(self,sql):
        self.__cur.execute(sql,)
        rows=self.__cur.fetchall()
        self.__conn.commit()
        return rows
    
    def InsertDB(self,sql):
        self.__cur.execute(sql,)
        self.__conn.commit()
        
    def DeleteDB(self,sql):
        self.__cur.execute(sql,)
        self.__conn.commit()
    
    def close(self):
        self.__conn.close()
        
#自动邮件通知函数
def Sendmail(table):
    admin_email = ['chuanguang_huo@dell.com']
    receivers = [table['email']]
    reason = table['reason']
    activity_id = table['break_fix_id']
    status = table['status']
    
    importlib.reload(sys)
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(receivers)
    sql_l2group_data = ''
    sql_l2group_data = table['team name'].strip()
    sql_l2group_list = []
    try:
        temp_L2Group_info = Dict_L2Group[sql_l2group_data]
        temp_L2Group_info_new = temp_L2Group_info.split(',')
        for temp_new in temp_L2Group_info_new:
            sql_l2group_list.append(temp_new)
    except KeyError:
        temp_L2Group_info = ''
        sql_l2group_list = ['chuanguang_huo@dell.com']
    
    mail.CC = ";".join(sql_l2group_list)
    mail.Subject = 'Japan car automation alert'
    if status == "error":
        mail.Body = "Please note!\nJapan car automation failure!\nFailure reasons:%s,\nPlease modify on delta and resubmit activity id on CDAT!\nActivity id:%s"%(reason,activity_id)
    elif status == 'manual' or status == '':
        if reason == 'unknown error' or reason == 'Code interrupt':
            #receivers.append(admin_email)
            #mail.To = ";".join(receivers)
            mail.To = ";".join(admin_email)
            mail.CC = ''
        mail.Body = "Please note!\nJapan car automation failure!\nFailure reasons:%s,please dispatch by yourself!\nActivity id:%s"%(reason,activity_id)
    elif status == 'amend':
        url_first="http://10.114.14.19:50/rework?Id="
        url=url_first+activity_id
        mail.Body = "Please note!\nJapan car automation failure!\nFailure reasons:%s,\nClick the link to rework,\nLink:%s"%(reason,url)
    elif status == 'review':
        url_first="http://10.114.14.19:50/review?Id="
        url=url_first+activity_id
        mail.Body = "Please note!\nJapan car automation pause!\nPause reasons:%s,\nBreak fix id:%s\nIf L2 has been approved,please click the link for confirmation,\nLink:%s"%(reason,activity_id,url)
    #mail.Attachments.Add('C:\Users\xxx\Desktop\git_auto_pull_new.py')
    #mail.To = ";".join(['chuanguang_huo@dell.com'])
    #mail.CC = ";".join(['chuanguang_huo@dell.com'])
    mail.Send()

def AgentDB(sql):
    conn=pymssql.connect(host=r"xmntsdb02",user=r"TCD_User",password=r"TCD_123",charset=r"utf8",database='Tool_CDAT')
    cur=conn.cursor()
    cur.execute(sql)
    rows=cur.fetchall()
    conn.close()
    return rows


def success_mail(table):
    receivers = [table['email']]
    importlib.reload(sys)
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(receivers)
    sql_l2group_data = ''
    sql_l2group_data = table['team name'].strip()
    sql_l2group_list = []
    try:
        temp_L2Group_info = Dict_L2Group[sql_l2group_data]
        temp_L2Group_info_new = temp_L2Group_info.split(',')
        for temp_new in temp_L2Group_info_new:
            sql_l2group_list.append(temp_new)
    except KeyError:
        temp_L2Group_info = ''
        sql_l2group_list = ['chuanguang_huo@dell.com']

    mail.CC = ";".join(sql_l2group_list)
    mail.Subject = 'Japan car automation notifications'
    mail.Body = "Please note!\nDispatch Successful!\nDPS:%s"%(table['DPS#'])
    mail.Send()

def error_sendmail(error_reason):
    importlib.reload(sys)
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'chuanguang_huo@dell.com'
    mail.Subject = 'Japan_car_error'
    mail.Body = "Please note!\nJapan car automation pause!\nPause reasons:%s"%error_reason
    mail.Send()


# =============================================================================
# table = {}
# table['email'] = 'chuanguang_huo@dell.com'
# table['team name'] = 'Test_Nt'
# table['DPS'] = 'dadsds'
# success_mail(table)
# =============================================================================
#==============================================================================
# wd = WriteClass()
# print(wd.write('80913523927','11','22','33'))
# wd.close()
#==============================================================================








