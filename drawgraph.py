#! /usr/bin/python
#coding:utf-8
import matplotlib
import datetime
import numpy as np
from numpy import arange
import matplotlib.pyplot as plt
from matplotlib import ticker
import matplotlib.dates as mdates
from matplotlib.dates import num2date, DateFormatter, DayLocator,HourLocator,SecondLocator, MinuteLocator,drange
from mpl_toolkits.axes_grid1 import make_axes_locatable
from mpl_toolkits.axes_grid.parasite_axes import SubplotHost

#######设置字体########
#cnfont = matplotlib.font_manager.FontProperties(fname='./STSONG.TTF')

class drawpic:
    '''用来生成图片'''
    def __init__(self,data=[],name='',dates=[],fontp=''):
        #self.sdate=sdate #起始日期
        #self.edate=edate #结束日期
        self.data=data # 做图数据列表集合
        self.name=name
        self.dates=dates
        self.cnfont = matplotlib.font_manager.FontProperties(fname=fontp+'STSONG.TTF')
        #self.start_list=self.sdate.split('-')
        #print self.start_list
        #self.end_list=self.edate.split('-') #由于drange的关系，需要取前一天日期
        #print self.end_list
        #self.k=datetime.datetime(int(self.end_list[0]),int(self.end_list[1]),int(self.end_list[2]))+ datetime.timedelta(1) 
        #self.k=datetime.datetime(int(self.end_list[0]),int(self.end_list[1]),int(self.end_list[2]))
        #print str(self.k)
        
    def _Gpic(self,xlabel,ylabel,title):
        '''xlabel 表示x轴的标签
           ylabel 表示y轴的标签
           start 起始日期
           end 结束日期
        '''

        #plt.figure(figsize=(8,2.5))
        plt.figure(figsize=(32,12))
        #设定x轴
        
        #date2 = datetime.datetime(self.k.year,self.k.month,self.k.day)
        #date1 = datetime.datetime(int(self.start_list[0]),int(self.start_list[1]),int(self.start_list[2]))
        #delta = datetime.timedelta(hours=1) #设置副刻度每小时间隔
        #dates=drange(date1,date2,delta) #x 轴
        #print self.dates
        ax = plt.gca()
        ax.plot_date(self.dates,self.data,linestyle='-',color="blue",marker='o')
        #ax.plot_date(dates,self.data,linestyle='-',color="blue",marker='o')
        
        
        ax.set_xlabel(xlabel) #设置坐标轴名称
        ax.set_ylabel(ylabel)
        ax.set_title(title,fontproperties=self.cnfont,fontsize=14)       
        plt.subplots_adjust(bottom = 0.15)
        #plt.subplots_adjust(bottom = 0.3)        
        #ax.yaxis.set_major_formatter( DateFormatter('%d') )
        #ax.xaxis.set_major_locator( DayLocator() )
        #ax.xaxis.set_minor_locator( HourLocator(arange(0,25,1)) )
        #ax.xaxis.set_major_locator( HourLocator(arange(0,24,1)) )
        # ax.xaxis.set_major_locator(MinuteLocator(range(5,60,5)))
        #ax.xaxis.set_major_formatter( DateFormatter('%d') )        
        ax.fmt_xdata = DateFormatter('%Y-%m-%d %H:%M:%S')        
        #plt.xticks(rotation=90)
        #plt.xticks(t, datestr)
        ax.xaxis.set_major_formatter(DateFormatter('%Y-%m-%d %H:%M:%S'))
        for label in ax.xaxis.get_ticklabels():
           label.set_rotation(45) #设置标签倾斜45度              
        plt.grid() #打开网格
        plt.show()   
        plt.savefig(self.name)
