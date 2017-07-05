#-*- coding:utf8 -*-
import os, sys
import time
import xlsxwriter as xlw


ReportDir = "./report/"
DataDir="./data/"

TimePre = time.strftime("%Y-%m-%d")
ReportName = TimePre + "-report.xlsx"
ReportPath = os.path.join(ReportDir,ReportName)

report_dic = {}

def Config_init():
    try :
        fconfig = open('province_name.ini','rb')
        for pro_line in fconfig :
            if pro_line.startswith('#') :
                continue
            try :
                CnName, AbbName, Ip, Company = pro_line.strip().split()
            except ValueError as ve:
                print "Error line: ",pro_line
                continue
            if not report_dic.has_key(AbbName):
                report_dic[AbbName] = {}
            else :
                print AbbName,"repeat!!"
                continue
            

        fconfig.close()
    except IOError as ioe :
        print ioe
        sys.exit(1)
        

def Load_log_system():
    pass

def Load_log_logstat():
    pass

def Load_log_intact():
    pass


def Create_excel():
    pass

def Write_excel():
    pass

def Report():
    pass

