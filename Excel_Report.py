#!/usr/bin/python
#-*- coding:utf8 -*-
import os, sys
import time
import xlsxwriter as xlw


ReportDir = "./report/"
DataDir="./data/"

TimePre = time.strftime("%Y-%m-%d")
ReportName = TimePre + "-report.xlsx"
ReportPath = os.path.join(ReportDir,ReportName)


SystemStatus = {}
LogStatus = {}
IntactStatus = {}

def load_log():
    BClog = TimePre + "-BClog.txt"
    IntactLog = TimePre + "-Intact.txt"
    SystemLog = TimePre + "-SystemStatus.txt"
    BClogPath = os.path.join(DataDir, BClog)
    IntactLogPath = os.path.join(DataDir, IntactLog)
    SystemLogPath = os.path.join(DataDir, SystemLog)
    if not os.path.exists(SystemLogPath) :
        print "NO such file: " + SystemLogPath
     
    fsystem = open(SystemLogPath,'rb')
    for record in fsystem :
        if record.startswith('#') :
            continue
        recordList = record.strip().split()
        ProvName = recordList[0]
        if not  SystemStatus.has_key(ProvName) :
            SystemStatus[ProvName] = {}
        SystemStatus[ProvName]['ip'] = recordList[1]
        SystemStatus[ProvName]['PStatus'] = recordList[2]
        SystemStatus[ProvName]['DBstatus'] = recordList[3]
        SystemStatus[ProvName]['DBDataStatus'] = recordList[4]
        SystemStatus[ProvName]['XMLStatus'] = recordList[5]
        SystemStatus[ProvName]['DBFileStatus'] = recordList[6]
     
    fsystem.close()
     
    if not os.path.exists(BClogPath):
        print "No such file: " + BClogPath
     
    fbclog = open(BClogPath,'rb')
    for record in fbclog :
        if record.startswith('#') :
            continue
        recordList = record.strip().split()
        ProvName, IspName = recordList[0], recordList[1]
        if not LogStatus.has_key(ProvName):
            LogStatus[ProvName] = {}
        if not LogStatus[ProvName].has_key(IspName):
            LogStatus[ProvName][IspName] = {}
        LogStatus[ProvName][IspName]['LogUp'] = recordList[2] 
        LogStatus[ProvName][IspName]['AccNum'] = recordList[3]
        LogStatus[ProvName][IspName]['liveNum'] = recordList[4]
        LogStatus[ProvName][IspName]['Src'] = recordList[5]
        LogStatus[ProvName][IspName]['DialIp'] = recordList[6]
    fbclog.close()
     
    fintact = open(IntactLogPath,'rb')
    for record in fintact:
        if record.startswith('#') :
            continue
        recordList = record.strip().split()
        ProvName, IspName = recordList[0], recordList[1]
        if not IntactStatus.has_key(ProvName):
            IntactStatus[ProvName] = {}
        if not IntactStatus[ProvName].has_key(IspName):
            IntactStatus[ProvName][IspName] = {}
        IntactStatus[ProvName][IspName]['LogNum'] = recordList[2] 
        IntactStatus[ProvName][IspName]['Intact'] = recordList[3]
    fintact.close()
load_log()
      
     
  
  
def Create_Header(SystemSheet, IntactSheet):
    header_format = workbook.add_format({
                                         'bold':1,
                                         'border':1,
                                         'align':'center',
                                         'valign':'vcenter',
                                         'fg_color':'8DB4E3'
                                         })
#设置系统运行状态的标题
    SystemSheet.set_column('A:A',10)
    SystemSheet.set_column('B:J',14)
    SystemSheet.merge_range('A1:A2', u'省份', header_format)
    SystemSheet.merge_range('B1:B2', u'服务器IP', header_format)
    SystemSheet.merge_range('C1:D1', u'服务器状态监控', header_format)
    SystemSheet.merge_range('E1:F1', u'数据库运行监控', header_format)
    SystemSheet.merge_range('G1:I1', u'程序运行监控', header_format)
    SystemSheet.merge_range('J1:J2', u'数据负责单位', header_format)
    SystemSheet.write('C2',u'网络情况',header_format )
    SystemSheet.write('D2',u'系统运行' ,header_format )
    SystemSheet.write('E2',u'运行状态' ,header_format )
    SystemSheet.write('F2',u'前一天数据' ,header_format )
    SystemSheet.write('G2',u'是否启动' ,header_format )
    SystemSheet.write('H2',u'数据库查询' ,header_format )
    SystemSheet.write('I2',u'探针文件' ,header_format )
      
#设置一致性比对结果的标题
    IntactSheet.set_column('A:A',10)
    IntactSheet.set_column('B:I',14)
    IntactSheet.merge_range('A1:A2', u'省份', header_format)
    IntactSheet.merge_range('B1:B2', u'运营商', header_format)
    IntactSheet.merge_range('C1:F1', u'拨号日志质量监控', header_format)
    IntactSheet.merge_range('G1:H1', u'一致性比对结果', header_format)
    IntactSheet.merge_range('I1:I2', u'数据负责单位', header_format)
    IntactSheet.write('C2',u'日志生成',header_format )
    IntactSheet.write('D2',u'日志条数',header_format )
    IntactSheet.write('E2',u'帐号总数',header_format )
    IntactSheet.write('F2',u'在线帐号数',header_format )
    IntactSheet.write('G2',u'比对日志数',header_format )
    IntactSheet.write('H2',u'数据完整率',header_format )
  
workbook = xlw.Workbook(ReportPath)
SystemSheet = workbook.add_worksheet(u"系统运行状态")
IntactSheet = workbook.add_worksheet(u"一致性比对结果") 
 
Create_Header(SystemSheet, IntactSheet)
Normal_format = workbook.add_format({'border':1,'align':'center','valign':'vcenter'})
Warn_format = workbook.add_format({'fg_color':'red','border':1,'align':'center','valign':'vcenter'})
NULL_format = workbook.add_format({'fg_color':'gray','border':1,'align':'center','valign':'vcenter'})
  
fconf = open("province_name.ini", 'rb')
IntactRow = 3
SystemRow = 3
for Prov in fconf :
    format = Normal_format
    if Prov.startswith('#') :
        continue
    CnName, AbbName, Ip, Company = Prov.strip().split()
    CnName = unicode(CnName,"utf-8")
    Company = unicode(Company,"utf-8")
    if SystemStatus.has_key(AbbName) :
        NetStatus = "YES"
        SysStatus = "YES"
        ip = SystemStatus[AbbName]['ip']
        PStatus = SystemStatus[AbbName]['PStatus']
        DBstatus = SystemStatus[AbbName]['DBstatus']
        DBDataStatus = SystemStatus[AbbName]['DBDataStatus']
        XMLStatus = SystemStatus[AbbName]['XMLStatus']
        DBFileStatus = SystemStatus[AbbName]['DBFileStatus']
        SystemSheet.write('A'+str(SystemRow),CnName,format)
        SystemSheet.write('B'+str(SystemRow),ip,format)
        SystemSheet.write('C'+str(SystemRow),NetStatus,format)
        SystemSheet.write('D'+str(SystemRow),NetStatus,format)

        if DBstatus == "NO" :
            format = Warn_format
        else :
            format = Normal_format
        SystemSheet.write('E'+str(SystemRow),DBstatus,format)
        if DBDataStatus == "NO" :
            format = Warn_format
        else :
            format = Normal_format
        SystemSheet.write('F'+str(SystemRow),DBDataStatus,format)
        if PStatus == "NO" :
            format = Warn_format
        else :
            format = Normal_format
        SystemSheet.write('G'+str(SystemRow),PStatus,format)       
        if DBFileStatus == "NO":
            format = Warn_format
        else :
            format = Normal_format
        SystemSheet.write('H'+str(SystemRow),DBFileStatus,format)
        if XMLStatus == "NO":
            format = Warn_format
        else :
            format = Normal_format
        SystemSheet.write('I'+str(SystemRow),XMLStatus,format)
        SystemSheet.write('J'+str(SystemRow),Company,Normal_format)
        SystemRow += 1
	format = Normal_format
# 完整性比对结果
    IntactSheet.merge_range('A'+str(IntactRow)+':A'+str(IntactRow+1),CnName,Normal_format )
    IntactSheet.write('B'+str(IntactRow),u'联通',Normal_format )
    IntactSheet.write('B'+str(IntactRow + 1),u'电信',Normal_format )
    IntactSheet.merge_range('I'+str(IntactRow)+':I'+str(IntactRow+1),Company,Normal_format )
    if IntactStatus.has_key(AbbName):
        
        if IntactStatus[AbbName].has_key('LT') :
            LogNum = IntactStatus[AbbName]['LT']['LogNum']
            Intact = IntactStatus[AbbName]['LT']['Intact']
            if LogNum == 0 :
                format = Warn_format
            else :
                format = Normal_format
            IntactSheet.write('G'+str(IntactRow),LogNum, format)
            if Intact == "0%" :
                format = Warn_format
            else :
                format = Normal_format
            IntactSheet.write('H' + str(IntactRow), Intact,format)
        else :
            IntactSheet.write('G' + str(IntactRow), "", NULL_format)
            IntactSheet.write('H' + str(IntactRow), "", NULL_format)
        if IntactStatus[AbbName].has_key('DX') :
            LogNum = IntactStatus[AbbName]['DX']['LogNum']
            Intact = IntactStatus[AbbName]['DX']['Intact']
            if LogNum == 0 :
                format = Warn_format
            else :
                format = Normal_format
            IntactSheet.write('G'+str(IntactRow+1),LogNum, format)
            if Intact == "0%" :
                format = Warn_format
            else :
                format = Normal_format
            IntactSheet.write('H' + str(IntactRow+1),Intact, format)
        else :
            IntactSheet.write('G' + str(IntactRow+1), "", NULL_format)
            IntactSheet.write('H' + str(IntactRow+1), "", NULL_format)            
           
    else :
        IntactSheet.write('G' + str(IntactRow), "", NULL_format)
        IntactSheet.write('H' + str(IntactRow), "", NULL_format)
        IntactSheet.write('G' + str(IntactRow+1), "", NULL_format)
        IntactSheet.write('H' + str(IntactRow+1), "", NULL_format)
    format = Normal_format
    if LogStatus.has_key(AbbName):

        if LogStatus[AbbName].has_key('LT'):
            LogUp = LogStatus[AbbName]['LT']['LogUp'] 
            AccNum = LogStatus[AbbName]['LT']['AccNum']
            liveNum = LogStatus[AbbName]['LT']['liveNum']
            Src = LogStatus[AbbName]['LT']['Src']
            DialIp = LogStatus[AbbName]['LT']['DialIp']
            IntactSheet.write('C' + str(IntactRow), Src, format)
            if LogUp == "YES" :
                format = Normal_format
            else :
                format = Warn_format
            IntactSheet.write('D' + str(IntactRow), LogUp, format)
            IntactSheet.write('E' + str(IntactRow), AccNum, Normal_format)
            if liveNum == '0' :
                format = Warn_format
            else :
                format = Normal_format
            IntactSheet.write('F' + str(IntactRow), liveNum, format)

        else :
            IntactSheet.write('C' + str(IntactRow), "", NULL_format)
            IntactSheet.write('D' + str(IntactRow), "", NULL_format)
            IntactSheet.write('E' + str(IntactRow), "", NULL_format)
            IntactSheet.write('F' + str(IntactRow), "", NULL_format)
            
        if LogStatus[AbbName].has_key('DX'):
            LogUp = LogStatus[AbbName]['DX']['LogUp'] 
            AccNum = LogStatus[AbbName]['DX']['AccNum']
            liveNum = LogStatus[AbbName]['DX']['liveNum']
            Src = LogStatus[AbbName]['DX']['Src']
            DialIp = LogStatus[AbbName]['DX']['DialIp']
            IntactSheet.write('C' + str(IntactRow+1), Src, format)
            if LogUp == "YES" :
                format = Normal_format
            else :
                format = Warn_format
            IntactSheet.write('D' + str(IntactRow+1), LogUp, format)
            IntactSheet.write('E' + str(IntactRow+1), AccNum, Normal_format)
            if liveNum == '0' :
                format = Warn_format
            else :
                format = Normal_format
            IntactSheet.write('F' + str(IntactRow+1), liveNum, format)

        else :
            IntactSheet.write('C' + str(IntactRow+1), "", NULL_format)
            IntactSheet.write('D' + str(IntactRow+1), "", NULL_format)
            IntactSheet.write('E' + str(IntactRow+1), "", NULL_format)
            IntactSheet.write('F' + str(IntactRow+1), "", NULL_format)
    else :
            IntactSheet.write('C' + str(IntactRow), "", NULL_format)
            IntactSheet.write('D' + str(IntactRow), "", NULL_format)
            IntactSheet.write('E' + str(IntactRow), "", NULL_format)
            IntactSheet.write('F' + str(IntactRow), "", NULL_format)        
            IntactSheet.write('C' + str(IntactRow+1), "", NULL_format)
            IntactSheet.write('D' + str(IntactRow+1), "", NULL_format)
            IntactSheet.write('E' + str(IntactRow+1), "", NULL_format)
            IntactSheet.write('F' + str(IntactRow+1), "", NULL_format)
    IntactRow += 2
    format = Normal_format
workbook.close()

    
