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
        print "No such file: " + SystemLogPath
     
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
#璁剧疆绯荤粺杩愯鐘舵�佺殑鏍囬
    SystemSheet.set_column('A:A',10)
    SystemSheet.set_column('B:J',14)
    SystemSheet.merge_range('A1:A2', u'鐪佷唤', header_format)
    SystemSheet.merge_range('B1:B2', u'鏈嶅姟鍣↖P', header_format)
    SystemSheet.merge_range('C1:D1', u'鏈嶅姟鍣ㄧ姸鎬佺洃鎺�', header_format)
    SystemSheet.merge_range('E1:F1', u'鏁版嵁搴撹繍琛岀洃鎺�', header_format)
    SystemSheet.merge_range('G1:I1', u'绋嬪簭杩愯鐩戞帶', header_format)
    SystemSheet.merge_range('J1:J2', u'鏁版嵁璐熻矗鍗曚綅', header_format)
    SystemSheet.write('C2',u'缃戠粶鎯呭喌',header_format )
    SystemSheet.write('D2',u'绯荤粺杩愯' ,header_format )
    SystemSheet.write('E2',u'杩愯鐘舵��' ,header_format )
    SystemSheet.write('F2',u'鍓嶄竴澶╂暟鎹�' ,header_format )
    SystemSheet.write('G2',u'鏄惁鍚姩' ,header_format )
    SystemSheet.write('H2',u'鏁版嵁搴撴煡璇�' ,header_format )
    SystemSheet.write('I2',u'鎺㈤拡鏂囦欢' ,header_format )
      
#璁剧疆涓�鑷存�ф瘮瀵圭粨鏋滅殑鏍囬
    IntactSheet.set_column('A:A',10)
    IntactSheet.set_column('B:H',14)
    IntactSheet.merge_range('A1:A2', u'鐪佷唤', header_format)
    IntactSheet.merge_range('B1:B2', u'杩愯惀鍟�', header_format)
    IntactSheet.merge_range('C1:F1', u'鎷ㄥ彿鏃ュ織璐ㄩ噺鐩戞帶', header_format)
    IntactSheet.merge_range('G1:H1', u'涓�鑷存�ф瘮瀵圭粨鏋�', header_format)
    IntactSheet.write('C2',u'鏃ュ織鐢熸垚',header_format )
    IntactSheet.write('D2',u'鏃ュ織鏉℃暟',header_format )
    IntactSheet.write('E2',u'甯愬彿鎬绘暟',header_format )
    IntactSheet.write('F2',u'鍦ㄧ嚎甯愬彿鏁�',header_format )
    IntactSheet.write('G2',u'姣斿鏃ュ織鏁�',header_format )
    IntactSheet.write('H2',u'鏁版嵁瀹屾暣鐜�',header_format )
  
workbook = xlw.Workbook(ReportPath)
SystemSheet = workbook.add_worksheet(u"绯荤粺杩愯鐘舵��")
IntactSheet = workbook.add_worksheet(u"涓�鑷存�ф瘮瀵圭粨鏋�") 
 
Create_Header(SystemSheet, IntactSheet)
Normal_format = workbook.add_format({'border':1})
Warn_format = workbook.add_format({'fg_color':'red','border':1})
NULL_format = workbook.add_format({'fg_color':'gray','border':1})
  
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
        NetStatus = "yes"
        SysStatus = "yes"
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

        if DBstatus == "no" :
            format = Warn_format
        else :
            format = Normal_format
        SystemSheet.write('E'+str(SystemRow),DBstatus,format)
        if DBDataStatus == "no" :
            format = Warn_format
        else :
            format = Normal_format
        SystemSheet.write('F'+str(SystemRow),DBDataStatus,format)
        if PStatus == "no" :
            format = Warn_format
        else :
            format = Normal_format
        SystemSheet.write('G'+str(SystemRow),PStatus,format)       
        if DBFileStatus == "no":
            format = Warn_format
        else :
            format = Normal_format
        SystemSheet.write('H'+str(SystemRow),DBFileStatus,format)
        if XMLStatus == "no":
            format = Warn_format
        else :
            format = Normal_format
        SystemSheet.write('I'+str(SystemRow),XMLStatus,format)
        SystemSheet.write('J'+str(SystemRow),Company,Normal_format)
        SystemRow += 1
        
#     if IntactStatus.has_key(AbbName):
#        if IntactStatus[AbbName].has_key('LT') :
           
        
  
workbook.close()

    


