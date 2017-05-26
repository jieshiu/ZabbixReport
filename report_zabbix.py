#!/usr/bin/python
# -*- coding: utf-8 -*-

import os, sys, string
import ConfigParser,getopt
import time
from datetime import datetime,timedelta

#import genpdf
from reportpdf import genpdf
import genexcel
import genmail


def helpshow():
	  print('''
	         This program can generate and send report of Zabbix via email with Excel & PDF.
	             
	         Options include: 
	         -t : Report time
	         -c filename : Send email report according to options in config file
	         -v : Prints the version number
	         -h : Display this help
	         Written in Python 2.7, by jieshiu''')
          sys.exit()
          return

def main(): 	
    stattime=''  
    try:  
        opts,args = getopt.getopt(sys.argv[1:],"t:c:")  
        if len(opts) < 1:           
          helpshow() 
        for op,value in opts:  
          if op == "-t":  
            stattime = value  
          elif op == "-c":  
            configpath = value              
          elif op == "-v":
          	print('Version 0.1') 
          	sys.exit()
          else:  
            helpshow()           
        #print(opts)  
        #print(args)  
    except getopt.GetoptError:  
        print("params are not defined well!")  
            
    if 'configpath' not in dir():  
       print("-c param is needed,config file path must define!")  
       sys.exit(1)   
    if not os.path.exists(configpath):  
       print("file : %s is not exists"%configpath)  
       sys.exit(1)  
    reload(sys)   
    sys.setdefaultencoding('utf8')  
        
    cf = ConfigParser.ConfigParser()
    cf.read(configpath)
    
    zabbix_url = cf.get("zabbix", "zabbix_url")
    report_template = cf.get("zabbix", "report_template")
    report_outpath = cf.get("zabbix", "report_out")
    outgraph= cf.get("zabbix", "outgraph")
    starttime= cf.get("zabbix", "starttime")
    server_list= cf.get("zabbix", "server_list")
    fontpath=cf.get("zabbix", "fontpath")
    
    mysql_host=cf.get("mysql", "host")
    mysql_usrname=cf.get("mysql", "usr")
    mysql_usrpwd=cf.get("mysql", "passwd")
    mysql_db=cf.get("mysql", "db")

    sendpdf=cf.get("email","sendpdf")
    smtp_server=cf.get("email","smtp_server")
    smtp_user=cf.get("email","smtp_user")
    smtp_pass=cf.get("email","smtp_pass")
    from_add=cf.get("email","from_add")
    to_add=cf.get("email","to_add").split(',')
    subject_c=cf.get("email","subject")
    htmlText=cf.get("email","htmlText")
    att_name=cf.get("email","att_name")    

    current_time=time.localtime()

    htmlfile_name = att_name+time.strftime("%Y%m%d%H%M%S",current_time)+r'.html'
    excelfile_name = att_name+time.strftime("%Y%m%d%H%M%S",current_time)+r'.xlsx'
    pdffile_name = att_name+time.strftime("%Y%m%d%H%M%S",current_time)+r'.pdf'
    
    if stattime != '':
    	 stattimelist=stattime.split(',')
    	 current_time = time.strptime(stattimelist[0], '%Y%m%d_%H_%M_%S')
    else:
    	 nowdate = datetime.now()
    	 date_str = nowdate.strftime("%Y%m%d")
    	 date_str=date_str+'_'+starttime
    	 tmpdt=datetime.strptime(date_str, "%Y%m%d_%H_%M_%S")    	
    	 tomorrow = tmpdt + timedelta(days=1)
    	 tomorrow_str = tomorrow.strftime("%Y%m%d_%H_%M_%S")
    	 stattime=date_str+','+tomorrow_str
    	        
    subject=subject_c+'_'+stattime
        
    pdfdata=genexcel.genexcel(outpath=report_outpath, excel_filename=excelfile_name, stattime=stattime,hostip=mysql_host,usrname=mysql_usrname,usrpsw=mysql_usrpwd,dbname=mysql_db,outgraph=outgraph,fontpath=fontpath,iplist=server_list)
    if outgraph=='yes':
       genpdf(current_time,pdffile_name,report_outpath,fontpath,pdfdata)

    authInfo={}
    authInfo['server'] = smtp_server
    authInfo['user'] = smtp_user
    authInfo['password'] = smtp_pass
    
    if sendpdf=='yes':
    	 genmail.sendEmail(authInfo, from_add, to_add, subject, '', htmlText, report_outpath,excelfile_name,pdffile_name)
    else:
       genmail.sendEmail(authInfo, from_add, to_add, subject, '', htmlText, report_outpath,excelfile_name,'')
    return

def ConvertCN(s):  
    return s.encode('utf8')
    
    
#if __name__ =="__main__":    
main()
