#/usr/bin/python
#coding:utf-8

import time
import datetime
import reportlab.rl_config
reportlab.rl_config.warnOnMissingFontGlyphs = 0
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen.canvas import Canvas
import calendar
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import fonts,colors
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate,Table,Image,PageTemplate,Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus.doctemplate import *
#from GerData import *
import copy

##注册中文字体
#pdfmetrics.registerFont(TTFont('song',"./STSONG.TTF")) #路径
#pdfmetrics.registerFont(TTFont('hei','./simhei.ttf')) #粗体
#fonts.addMapping('song',0,0,'song')
#fonts.addMapping('song',0,1,'song')
#fonts.addMapping('song',1,0,'hei')
#fonts.addMapping('song',1,1,'hei')

class MyPDFdoc:
    class MyPageTemp(PageTemplate):
        '''
        定义一个页面模板
        '''
        def __init__(self):
            F6=Frame(x1=0.5*inch, y1=0.5*inch, width=7.5*inch, height=2.0*inch,showBoundary=0)
            F5=Frame(x1=0.5*inch, y1=2.5*inch, width=7.5*inch, height=2.0*inch,showBoundary=0)
            F4=Frame(x1=0.5*inch, y1=4.5*inch, width=7.5*inch, height=2.0*inch,showBoundary=0)
            F3=Frame(x1=0.5*inch, y1=6.5*inch, width=7.5*inch, height=2.0*inch,showBoundary=0)
            F2=Frame(x1=0.5*inch, y1=8.5*inch, width=7.5*inch, height=2.0*inch,showBoundary=0)
            F1=Frame(x1=0.5*inch, y1=10.5*inch, width=7.5*inch, height=0.5*inch,showBoundary=0)
            PageTemplate.__init__(self,"MyTemplate",[F1,F2,F3,F4,F5,F6])

        def beforeDrawPage(self,canvas,doc): #在页面生成之前做什么画logo
            pass

    def __init__(self,filename,fontpath):    	  
    	self.fontpath=fontpath+'STSONG.TTF'    	  
    	pdfmetrics.registerFont(TTFont('song',self.fontpath)) #路径
        pdfmetrics.registerFont(TTFont('hei',fontpath+'simhei.ttf')) #粗体
        fonts.addMapping('song',0,0,'song')
        fonts.addMapping('song',0,1,'song')
        fonts.addMapping('song',1,0,'hei')
        fonts.addMapping('song',1,1,'hei')
        self.filename=filename
        self.objects=[] ##story
        self.doc=BaseDocTemplate(self.filename)
        self.doc.addPageTemplates(self.MyPageTemp())
        self.Style=getSampleStyleSheet()
        #设置中文,设置四种格式
        #self.nor=self.Style['Normal']
        self.cn=self.Style['Normal']
        self.cn.fontName='song'
        self.cn.fontSize=9
        self.t=self.Style['Title']
        self.t.fontName='song'
        self.t.fontSize=15
        self.h=self.Style['Heading1']
        self.h.fontName='song'
        self.h.fontSize=10
        self.end=copy.deepcopy(self.t)
        self.end.fontSize=7

    def _CreatePDF(self,timestr,dep,pro,data): #        
        d=[]
        data_1=[]
        tmp_arr=[]
        col=[u'设备IP',u'cpu使用率均值',u'cpu使用率峰值',u'最小内存空闲GB',u'磁盘idle均值',u'最小磁盘idle']
        '''
        遍历字典       
        '''
        #生成table Style
        S1=[('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),('BOX', (0,0), (-1,-1), 0.25, colors.black)]
        S2=[('INNERGRID', (0,0), (-1,-1), 0.4, colors.black),('BOX', (0,0), (-1,-1), 0.25, colors.black),('ALIGN',(0,0),(-1,-1),'CENTER')]
        
        #添加标题
        self.objects.append(Paragraph('''<b><u>%s%s</u></b>%sServer Report''' % (dep,pro,timestr),self.t))
        self.objects.append(FrameBreak()) #切换到第二个FORM
        self.objects.append(Paragraph('''<b>Performance Profile</b>''',self.h))
        if data.has_key('table'):            
            for y in col:
                tmp_arr.append(Paragraph('''<b>%s</b>''' % y,self.cn))
            for i in data['table']:
                t0=Paragraph('%s'%(i),self.cn)
                d=data['table'][i]
                t1=Paragraph('%s'%(d[0]),self.cn)
                t2=Paragraph('%s'%(d[1]),self.cn)
                t3=Paragraph('%s'%(d[2]),self.cn)
                t4=Paragraph('%s'%(d[3]),self.cn)
                t5=Paragraph('%s'%(d[4]),self.cn)                
                data_1.append([t0,t1,t2,t3,t4,t5])               
            data_1.insert(0,tmp_arr)
            t=Table(data_1,style=S1,colWidths=inch,rowHeights=0.2*inch)            
            self.objects.append(t)
        if data.has_key('pic'): 
           for i in data['pic']:        	    
               self.objects.append(FrameBreak())
               self.objects.append(Paragraph('''<b>%s性能指标趋势图</b>'''%(i),self.h))
               self.objects.append(Spacer(0,0))        
               self.objects.append(Image('%s' % (data['pic'][i]),500,102))       
        #self.objects.append(FrameBreak())        
        self.objects.append(Spacer(0.1*inch,0.1*inch))
        self.objects.append(Paragraph('''<font color=red>The above data is for reference only and cannot be used as a final measure！</font>''',self.end))
        self.doc.build(self.objects)

def genpdf(stime,pdf_name,out_path,fontpath,pdfdata):
    timestr=time.strftime("%Y年%m月%d日",stime)    
    pdf=MyPDFdoc(out_path+pdf_name,fontpath)
    pdf._CreatePDF(timestr,'xxxx','xx',pdfdata)
    return
    
