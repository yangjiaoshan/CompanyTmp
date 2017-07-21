#-*- coding:utf8 -*-
import os, sys
import time
import xlsxwriter as xlw


ReportDir = "./report/"
DataDir="./data/"

TimePre = time.strftime("%Y-%m-%d")
TimePre = '2017-07-16'
ReportName = TimePre + "-report.xlsx"
ReportPath = os.path.join(ReportDir,ReportName)

report_dic = {}
pro_list = []

def Config_init():
    try :
        fconfig = open('province_name.ini','rb')
        for pro_line in fconfig :
            if pro_line.startswith('#') :
                continue
            try :
                CnName, AbbName, Ip, Company = pro_line.strip().split()
                CnName = unicode(CnName,"utf-8")
                Company = unicode(Company,"utf-8")
                pro_list.append(AbbName)
            except ValueError as ve: 
                print ve
                print "Error line: ",pro_line
                continue
            if not report_dic.has_key(AbbName):
                report_dic[AbbName] = {}
            else :
                print AbbName,"repeat!!"
                continue
            report_dic[AbbName]['system'] = {}
            report_dic[AbbName]['logstat'] = {}
            report_dic[AbbName]['intact'] = {}
            report_dic[AbbName]['company'] = Company
            report_dic[AbbName]['bc_ip'] = Ip
            report_dic[AbbName]['CnName'] = CnName
            report_dic[AbbName]['system']['net'] = ['NO',1]
            report_dic[AbbName]['system']['pro'] = ['null',2]
            report_dic[AbbName]['system']['dbstat'] = ['null',2]
            report_dic[AbbName]['system']['dbdata'] = ['null',2]
            report_dic[AbbName]['system']['dbfile'] = ['null',2]
            report_dic[AbbName]['system']['xml'] = ['null',2]
            report_dic[AbbName]['logstat']['dx'] = {}
            report_dic[AbbName]['logstat']['lt'] = {}
            report_dic[AbbName]['logstat']['dx']['ppp_ip'] = ['null',2]
            report_dic[AbbName]['logstat']['dx']['upload'] = ['null',2]
            report_dic[AbbName]['logstat']['dx']['acc_sum'] = ['null',2]
            report_dic[AbbName]['logstat']['dx']['ali_sum'] = ['null',2]
            report_dic[AbbName]['logstat']['dx']['src'] = ['null',2]
            report_dic[AbbName]['logstat']['lt']['ppp_ip'] = ['null',2]
            report_dic[AbbName]['logstat']['lt']['upload'] = ['null',2]
            report_dic[AbbName]['logstat']['lt']['acc_sum'] = ['null',2]
            report_dic[AbbName]['logstat']['lt']['ali_sum'] = ['null',2]
            report_dic[AbbName]['logstat']['lt']['src'] = ['null',2]
            report_dic[AbbName]['intact']['lt'] = {}
            report_dic[AbbName]['intact']['dx'] = {}
            report_dic[AbbName]['intact']['lt']['record'] = ['null',2]
            report_dic[AbbName]['intact']['lt']['intact'] = ['null',2]
            report_dic[AbbName]['intact']['dx']['record'] = ['null',2]
            report_dic[AbbName]['intact']['dx']['intact'] = ['null',2]
        fconfig.close()
    except IOError as ioe :
        print ioe
        sys.exit(1)
        

def Load_log_system():
    SystemLog = TimePre + "-SystemStatus.txt"
    SystemLogPath = os.path.join(DataDir, SystemLog)
    try :
        fsystem = open(SystemLogPath,'rb')
        for record in fsystem :
            if record.startswith('#') :
                continue
            recordList = record.strip().split()
            if len(recordList) != 7 :
                print "Colunm is fasle : ",record
                continue
            AbbName, Ip, Ps, DBs, DBd, xml, DBf =  recordList
            report_dic[AbbName]['system']['net'] = ['YES',0]
            if Ps.upper() == 'NO' :
                report_dic[AbbName]['system']['pro'] = ['NO',1]
            elif Ps.upper() == 'YES' :
                report_dic[AbbName]['system']['pro'] = ['YES',0]
                
            if DBs.upper() == 'NO' :
                report_dic[AbbName]['system']['dbstat'] = ['NO',1]
            elif DBs.upper() == 'YES' :
                report_dic[AbbName]['system']['dbstat'] = ['YES',0]
                            
            if DBd.upper() == 'NO' :
                report_dic[AbbName]['system']['dbdata'] = ['NO',1]
            elif DBd.upper() == 'YES' :
                report_dic[AbbName]['system']['dbdata'] = ['YES',0]
                
            if xml.upper() == 'NO' :
                report_dic[AbbName]['system']['xml'] = ['NO',1]
            elif xml.upper() == 'YES' :
                report_dic[AbbName]['system']['xml'] = ['YES',0]                

            if DBf.upper() == 'NO' :
                report_dic[AbbName]['system']['dbfile'] = ['NO',1]
            elif DBf.upper() == 'YES' :
                report_dic[AbbName]['system']['dbfile'] = ['YES',0]                            
        fsystem.close()
    except IOError as ioe :
        print ioe

def Load_log_logstat():
    BClog = TimePre + "-BClog.txt"
    BClogPath = os.path.join(DataDir, BClog)
    try :
        fbclog = open(BClogPath,'rb')
        for record in fbclog :
            if record.startswith('#') :
                continue
            recordList = record.strip().split()
            if len(recordList) != 7 :
                print "Colunm is fasle : ",record
                continue
            AbbName, Isp, Ups, Accn, Alin, src, ppp_ip =  recordList
            Isp = Isp.lower()
            
            if Ups.upper() == 'NO' :
                report_dic[AbbName]['logstat'][Isp]['upload'] = ['NO',1]
            elif Ups.upper() == 'YES' :
                report_dic[AbbName]['logstat'][Isp]['upload'] = ['YES',0]
                
            if Accn.upper() != 'NULL' :
                if Accn == '0' :
                    report_dic[AbbName]['logstat'][Isp]['acc_sum'] = [Accn,1]
                else :
                    report_dic[AbbName]['logstat'][Isp]['acc_sum'] = [Accn,0]
            if Alin.upper() != 'NULL' :                
                if Alin == '0' :
                    report_dic[AbbName]['logstat'][Isp]['ali_sum'] = [Alin,1]
                else :
                    report_dic[AbbName]['logstat'][Isp]['ali_sum'] = [Alin,0]

            report_dic[AbbName]['logstat'][Isp]['src'] = [src,0]
            report_dic[AbbName]['logstat'][Isp]['ppp_ip'] = [ppp_ip,0]
                
        fbclog.close()
    except IOError as ioe :
        print ioe
    
def Load_log_intact():
    IntactLog = TimePre + "-Intact.txt"
    IntactLogPath = os.path.join(DataDir, IntactLog)
    try :
        finlog = open(IntactLogPath,'rb')
        for record in finlog :
            if record.startswith('#') :
                continue
            recordList = record.strip().split()
            if len(recordList) != 4 :
                print "Colunm is fasle : ",record
                continue
            AbbName, Isp, rec_sum, intact =  recordList
            Isp = Isp.lower()
            if rec_sum.upper() == '0' :
                report_dic[AbbName]['intact'][Isp]['record'] = [rec_sum,1]
            else :
                report_dic[AbbName]['intact'][Isp]['record'] = [rec_sum,0]
                
            if intact.upper() != 'NULL' :   
                if intact == '0.00%' :
                    report_dic[AbbName]['intact'][Isp]['intact'] = [intact,1]
                    report_dic[AbbName]['intact'][Isp]['record'] = [rec_sum,1]
                else :
                    report_dic[AbbName]['intact'][Isp]['intact'] = [intact,0]
                            
                
        finlog.close()
    except IOError as ioe :
        print ioe

def Create_excel(workbook,SystemSheet,IntactSheet):
    header_format = workbook.add_format({
                                         'bold':1,
                                         'border':1,
                                         'align':'center',
                                         'valign':'vcenter',
                                         'fg_color':'8DB4E3'
                                         })
    SystemSheet.set_column('A:C',14)
    SystemSheet.set_column('D:J',14)
    SystemSheet.merge_range('A1:A2', u'省份', header_format)
    SystemSheet.merge_range('B1:B2', u'服务器IP', header_format)
    SystemSheet.merge_range('C1:C2', u'网络状态', header_format)
    SystemSheet.merge_range('D1:E1', u'数据库运行状态', header_format)
    SystemSheet.merge_range('F1:H1', u'程序运行监控', header_format)
    SystemSheet.merge_range('I1:I2', u'所属公司', header_format)
    SystemSheet.write('D2',u'运行状态' ,header_format )
    SystemSheet.write('E2',u'前一天数据' ,header_format )
    SystemSheet.write('F2',u'是否启动' ,header_format )
    SystemSheet.write('G2',u'数据库查询' ,header_format )
    SystemSheet.write('H2',u'探针文件' ,header_format )
      
    IntactSheet.set_column('A:B',10)
    IntactSheet.set_column('C:H',14)
    IntactSheet.merge_range('A1:A2', u'省份', header_format)
    IntactSheet.merge_range('B1:B2', u'运营商', header_format)
    IntactSheet.merge_range('C1:F1', u'拨号日志质量监控', header_format)
    IntactSheet.merge_range('G1:H1', u'一致性比对结果', header_format)
    IntactSheet.merge_range('I1:I2', u'所属公司', header_format)
    IntactSheet.write('C2',u'拨号日志生成',header_format )
    IntactSheet.write('D2',u'日志条数',header_format )
    IntactSheet.write('E2',u'账号总数',header_format )
    IntactSheet.write('F2',u'在线账号数',header_format )
    IntactSheet.write('G2',u'比对日志总数',header_format )
    IntactSheet.write('H2',u'数据完整率',header_format )
 
def Write_excel(workbook,sheet,site,value,state,mode):
    """
    sheet:表  ; site:表中位置 ; value:值 ; state:值对应的系统状态0正常，1告警，2为空  ;
    mode:single代表在一个单元格中，merge代表合并单元格
    """
    Normal_format = workbook.add_format({'border':1,'align':'center','valign':'vcenter'})
    Warn_format = workbook.add_format({'fg_color':'red','border':1,'align':'center','valign':'vcenter'})
    NULL_format = workbook.add_format({'fg_color':'gray','border':1,'align':'center','valign':'vcenter'})
    format = Normal_format
    if state == 1 :
        format = Warn_format
    elif state == 2 :
        format = NULL_format
    
    if mode == 'single':
        sheet.write(site,value,format)
    elif mode == "merge" :
        sheet.merge_range(site,value,format)
    
def Report():
    workbook = xlw.Workbook(ReportPath)
    SystemSheet = workbook.add_worksheet(u"系统运行")
    IntactSheet = workbook.add_worksheet(u"一致性比对")
    Create_excel(workbook,SystemSheet,IntactSheet)
    
    sys_index = 3
    intact_index = 3
    for AbbName in pro_list :
        pro_dic = report_dic[AbbName]
        Ip = pro_dic['bc_ip']
        Company = pro_dic['company']
        CnName = pro_dic['CnName']
        Write_excel(workbook,SystemSheet,'A' + str(sys_index),CnName, 0,'single')
        Write_excel(workbook,SystemSheet,'B' + str(sys_index),Ip, 0,'single')
        Write_excel(workbook,SystemSheet,'I' + str(sys_index),Company,0,'single')
        Write_excel(workbook,IntactSheet,'A' + str(intact_index) +':A' +str(intact_index + 1 ),CnName, 0,'merge')
        Write_excel(workbook,IntactSheet,'I' + str(intact_index) +':I' +str(intact_index + 1 ),Company, 0,'merge')
        Write_excel(workbook,IntactSheet,'B' + str(intact_index),u'联通', 0,'single')
        Write_excel(workbook,IntactSheet,'B' + str(intact_index + 1),u'电信', 0,'single')

        net_v, net_s = pro_dic['system']['net']
        pro_v, pro_s = pro_dic['system']['pro']
        dbs_v, dbs_s = pro_dic['system']['dbstat']
        dbd_v, dbd_s = pro_dic['system']['dbdata']
        dbf_v, dbf_s = pro_dic['system']['dbfile']
        xml_v, xml_s = pro_dic['system']['xml'] 
        Write_excel(workbook,SystemSheet,'C' + str(sys_index),net_v, net_s,'single')
        Write_excel(workbook,SystemSheet,'D' + str(sys_index),dbs_v, dbs_s,'single')
        Write_excel(workbook,SystemSheet,'E' + str(sys_index),dbd_v, dbd_s,'single')
        Write_excel(workbook,SystemSheet,'F' + str(sys_index),pro_v, pro_s,'single')
        Write_excel(workbook,SystemSheet,'G' + str(sys_index),dbf_v, dbf_s,'single')
        Write_excel(workbook,SystemSheet,'H' + str(sys_index),xml_v, xml_s,'single')
		
        lt_ip_v, lt_ip_s  = pro_dic['logstat']['lt']['ppp_ip'] 
        lt_up_v, lt_up_s  = pro_dic['logstat']['lt']['upload'] 
        lt_accn_v, lt_accn_s = pro_dic['logstat']['lt']['acc_sum']
        lt_alin_v, lt_alin_s = pro_dic['logstat']['lt']['ali_sum']
        lt_src_v, lt_src_s = pro_dic['logstat']['lt']['src']
        lt_rec_v, lt_rec_s = pro_dic['intact']['lt']['record']
        lt_int_v, lt_int_s = pro_dic['intact']['lt']['intact']

        dx_ip_v, dx_ip_s  = pro_dic['logstat']['dx']['ppp_ip'] 
        dx_up_v, dx_up_s  = pro_dic['logstat']['dx']['upload'] 
        dx_accn_v, dx_accn_s = pro_dic['logstat']['dx']['acc_sum']
        dx_alin_v, dx_alin_s = pro_dic['logstat']['dx']['ali_sum']
        dx_src_v, dx_src_s = pro_dic['logstat']['dx']['src']
        dx_rec_v, dx_rec_s = pro_dic['intact']['dx']['record']
        dx_int_v, dx_int_s = pro_dic['intact']['dx']['intact']
		
        Write_excel(workbook,IntactSheet,'C' + str(intact_index),lt_up_v, lt_up_s,'single')
        Write_excel(workbook,IntactSheet,'D' + str(intact_index),'NULL', 2,'single')
        Write_excel(workbook,IntactSheet,'E' + str(intact_index),lt_accn_v, lt_accn_s,'single')
        Write_excel(workbook,IntactSheet,'F' + str(intact_index),lt_alin_v, lt_alin_s,'single')
        Write_excel(workbook,IntactSheet,'G' + str(intact_index),lt_rec_v, lt_rec_s,'single')
        Write_excel(workbook,IntactSheet,'H' + str(intact_index),lt_int_v, lt_int_s,'single')

        Write_excel(workbook,IntactSheet,'C' + str(intact_index+1),dx_up_v, dx_up_s,'single')
        Write_excel(workbook,IntactSheet,'D' + str(intact_index+1),'NULL', 2,'single')
        Write_excel(workbook,IntactSheet,'E' + str(intact_index+1),dx_accn_v, dx_accn_s,'single')
        Write_excel(workbook,IntactSheet,'F' + str(intact_index+1),dx_alin_v, dx_alin_s,'single')
        Write_excel(workbook,IntactSheet,'G' + str(intact_index+1),dx_rec_v, dx_rec_s,'single')
        Write_excel(workbook,IntactSheet,'H' + str(intact_index+1),dx_int_v, dx_int_s,'single')

        sys_index += 1
        intact_index += 2
    workbook.close()                                             
    
    
if __name__ == "__main__" :
    Config_init()
    Load_log_system()
    Load_log_logstat()
    Load_log_intact()
#     print report_dic
    Report()
