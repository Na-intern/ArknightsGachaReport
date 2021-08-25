# -*- coding: utf-8 -*-
'''
created by 三鲸 21-08-25
写在源码中的备忘:pyqt5图形界面中不含QtWebEngineWidgets，因此要在ui文件中手动做出以下修改
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl
self.webBrowser = QWebEngineView(Arknights)

'''
import sys,os,json,re,requests,time
from win32com.client import Dispatch,GetObject
import pandas as pd

from PyQt5.QtCore import QFileInfo, QUrl
from PyQt5.QtGui import QIcon, QIntValidator
from PyQt5.QtWidgets import QAbstractItemView, QApplication, QFileDialog, QTableWidgetItem,QWidget,QDesktopWidget

from ui_Classify import *
from ui_GenReport import *
from ui_NoteSearch import *
from ui_RemoveNote import *
import Ark_icon
#变量说明
'''
df_base:本地读取的DataFrame数据
df_new:本次更新的DataFrame数据
df：汇总的DataFrame数据
time_set:集合，记录所有抽卡数据的时间
'''
path=os.getcwd()#路径
agent_search={'全部寻访':{3:0,4:0,5:0,6:0},'标准寻访':{3:0,4:0,5:0,6:0}}#字典，进行每类寻访各星级的计数
agent_search_chars={'全部寻访':[],'标注寻访':[]}#字典，记录每类寻访中的六星数据
agent_search_chars_5star={'全部寻访':[],'标注寻访':[]}#字典，记录每类寻访中的五星数据
agent_search_cnt={'全部寻访':1,'标注寻访':1}#字典，记录每类寻访中的六星数据
agent_search_cnt_5star={'全部寻访':1,'标注寻访':1}#字典，记录每类寻访中的五星数据
df_columns=['时间','干员','星级','寻访类别','所属寻访类别中累计未出六星次数','所属寻访类别中累计未出五星次数','全部寻访累计未出六星次数','全部寻访累计未出五星次数',]
cache=path+'\\UserData\\cache.txt'#用户信息缓存文件
try:
    with open(cache,'r',encoding='utf-8') as f:
        username,url=f.readlines()
        username,url=username.rstrip('\n'),url.rstrip('\n')
except FileNotFoundError:
    username,url='',''
def get_json(url,i):#从网页获取json格式抽卡数据并转化为字典
    url=re.sub("page=(.*?)&", 'page='+str(i)+'&',url)#用i替换url中的页码
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'}
    res= requests.get(url, headers=headers,timeout=3)
    res.encoding='utf-8'
    res=res.text
    res=json.loads(res)
    return res
def std_time(timestamp):#把时间戳转化成具体的年月日
    time_local = time.localtime(timestamp)
    dt = time.strftime("%Y-%m-%d %H:%M:%S",time_local)
    return dt
def solve_js(js):#从json格式抽卡数据转化得到的字典提取数据到df文件
    for i in range(len(js['data']['list'])):
        time=std_time(js['data']['list'][i]['ts'])
        if time not in time_set:#如果这个时间点的记录没有保存
            time_set.add(time)
            chars=list(js['data']['list'][i]['chars'])
            if len(chars)==10:#区分十连抽卡和单抽
                chars=chars[::-1]
            for char in chars:#越先抽到的卡放在数据的越上层
                rarity= char['rarity']+1
                df_new.columns=df_base.columns
                df_new.loc[df_new.shape[0]+1]=[time,char['name'],rarity,'标准寻访',0,0,0,0]#新增数据

class ArknightsGachaReportWindow(QWidget, Ui_Arknights):
    def __init__(self, parent=None):    
        super(ArknightsGachaReportWindow, self).__init__(parent)
        self.setupUi(self)
        self.retranslateUi(self)
        self.setWindowIcon(QIcon(':Ark.ico'))#图标设置
        self.setWindowTitle("明日方舟抽卡记录分析")#标题设置
        self.resize(1080,720)
        self.center()
        self.url_lineEdit.setText(url)
        self.username_lineEdit.setText(username)
        if os.path.exists("./UserData/ArknightsGachaReport.html"):
            self.webBrowser.load(QUrl(QFileInfo("./UserData/ArknightsGachaReport.html").absoluteFilePath()))
        else:
            self.webBrowser.load(QUrl("https://ak.hypergryph.com/index"))
    #自定义center函数     
    def center(self):  
        #描述屏幕#获得屏幕大小
        screen = QDesktopWidget().screenGeometry()  
        size = self.geometry()    
        self.move((screen.width() - size.width()) // 2,  (screen.height() - size.height()) // 2)  
    #三个槽函数
    def RenewData(self):
        global df,time_set,df_base,df_new,username,file_name
        #从输入中获取url和username并写入缓存
        url=self.url_lineEdit.text()
        username=self.username_lineEdit.text()
        if not os.path.isdir(path+'\\UserData'):
            os.mkdir(path+'\\UserData')
        with open(cache,'w+',encoding='utf-8') as f:
            f.write(username+'\n')
            f.write(url+'\n')
        try:
            #尝试从本地读取数据，否则新建
            local_valid=False
            self.PromptBox.setText( '……正在尝试从本地读取数据……')
            gui=QApplication.processEvents
            gui()
            file_name=path+'\\UserData\\'+username+'明日方舟抽卡记录.xlsx'
            try:
                df_base=pd.read_excel(file_name)
                df_base.columns=df_columns
                local_valid=True
            except:
                df_base=pd.DataFrame(columns=df_columns)
            del file_name#暂时删掉file_name，拿完数据后再复原，否则excel导出会出错
            #设置时间集合
            time_set={ts for ts in df_base['时间']}
            #寻访类别计入字典
            for s_name in set(df_base['寻访类别']):
                if s_name not in agent_search:
                    agent_search[s_name]={3:0,4:0,5:0,6:0}
            #初始化df文件
            df_new=pd.DataFrame(columns=df_base.columns)                        
            #获得json格式抽卡记录后对df数据文件进行修改
            for i in range(1,10+1):
                if local_valid==True:
                    self.PromptBox.setText('……本地读取数据成功，正在尝试从网络更新第%d页数据……'%i)
                else:
                    self.PromptBox.setText('……本地不存在"{}"的数据，正在尝试从网络获取第{}页数据……'.format(username,str(i)))
                gui=QApplication.processEvents
                gui()
                js=get_json(url,i)
                solve_js(js)
            df_new=df_new.iloc[::-1]
            df=pd.concat([df_base,df_new])
            df.index=[i for i in range(1,df.shape[0]+1)]#设置索引从0开始
            self.PromptBox.setText( '数据更新完毕')
            file_name=path+'\\UserData\\'+username+'明日方舟抽卡记录.xlsx'
            gui=QApplication.processEvents
            gui()
            ArknightsGachaReport.GenReport()
        except (json.JSONDecodeError,ConnectionError,KeyError):
            self.PromptBox.setText('数据更新出错，您输入的数据链接指向该页面，请检查链接是否出错。。。')
            self.webBrowser.load(QUrl(url))
            gui=QApplication.processEvents
            gui()
        except Exception as e:
            self.PromptBox.setText('报错：'+str(e))
            gui=QApplication.processEvents
            gui()
    def ViewData(self):
        try:
            type(df)
        except NameError:
            self.PromptBox.setText( '[获取数据]后才能执行[寻访分类]。。。')
            gui=QApplication.processEvents
            gui()
            pass
        else:
            self.DataWindow =GachaDataView()  
            self.DataWindow.show()  
        return
    def ExtractExcel(self):
        try:
            type(file_name)
        except NameError:
            self.PromptBox.setText( '[获取数据]后才能执行[导出excel]。。。')
            gui=QApplication.processEvents
            gui()
        else:
            fname,_=QFileDialog().getSaveFileName(self,'Save',path+'\\{}明日方舟抽卡记录.xlsx'.format(username),"Excel文件(*.xlsx)")
            #抽卡记录美化
            import xlwings as xw
            gacha_color={1:0x9c9c9c,2:0x9c9c9c,3:0xcd944f,4:0x8b1a55,5:0x698cff,6:0x3030ff,'1':0x9c9c9c,'2':0x9c9c9c,'3':0xcd944f,'4':0x8b1a55,'5':0x698cff,'6':0x3030ff}#根据抽到的不同星级的卡设置颜色#用BGR16位表示
            self.PromptBox.setText('……正在准备导出excel……')
            gui=QApplication.processEvents
            gui()
            try:
                app=xw.App(visible=False,add_book=False)
                wb=app.books.open(file_name)
                #写入全部寻访数据
                sht = wb.sheets[0]
                sht.name='全部寻访'
                wb.app.api.ActiveWindow.SplitRow = 1  # 冻结第一行
                wb.app.api.FreezePanes = True
                info = sht.used_range
                nrows = info.last_cell.row
                rng=re.findall('!(.*?)>',str(info))[0]#从info里提取得到整个工作簿的有效范围
                rng=xw.Range(rng)
                rng.autofit()#设置单元格大小自适应
                rng.api.Font.Name = '黑体'#设置字体为黑体
                for i in range(1,nrows+1):
                    self.PromptBox.setText( '……正在预处理[全部寻访]数据，进度{:.1%}……'.format(i/(nrows+1)))
                    gui=QApplication.processEvents
                    gui()
                    rarity=sht.range(i,3).value
                    if rarity in gacha_color:
                        sht.range(i,1).expand('right').api.Font.Color =gacha_color[rarity]
                #对单组寻访进行设置
                # for s_name in agent_search.keys():
                #     if s_name!='全部寻访':
                #         sht = wb.sheets.add(s_name)#命名
                #         wb.app.api.ActiveWindow.SplitRow = 1  # 冻结第一行
                #         wb.app.api.FreezePanes = True
                #         row_cnt=1#row_cnt用来指示行位置
                #         sht.range(row_cnt,1).expand('right').value = list(df.columns)[:6]
                #         for i in range(1,df.shape[0]+1):
                #             self.PromptBox.setText( '……正在写入:[{}]数据，进度{:.1%}……'.format(s_name,i/(nrows+1)))
                #             gui=QApplication.processEvents
                #             gui()
                #             if df.loc[i]['寻访类别']==s_name:
                #                 row_cnt+=1
                #                 sht.range(row_cnt,1).expand('right').value = df.loc[i].tolist()[:6]
                #                 rarity=sht.range(i,3).value
                #                 if rarity in gacha_color:
                #                     sht.range(row_cnt,1).expand('right').api.Font.Color =gacha_color[rarity]
                #         info = sht.used_range
                #         rng=re.findall('!(.*?)>',str(info))[0]#从info里提取得到整个工作簿的有效范围
                #         rng=xw.Range(rng)
                #         rng.autofit()#设置单元格大小自适应
                #         rng.api.Font.Name = '黑体'#设置字体为黑体
                wb.save(fname)
                app.kill()
                try:
                    excel_app=Dispatch('Excel.Application')
                    self.PromptBox.setText( '……正在尝试设置筛选标签，稍后可能需要手动保存更改……')
                    gui=QApplication.processEvents
                    gui()
                    excel_app.Workbooks.Open(os.path.abspath(fname))
                    excel_app.Selection.AutoFilter(Field=1)
                    excel_app.Quit()
                    self.PromptBox.setText( 'excel文件已导出,请点击弹出的对话框保存更改')
                    gui=QApplication.processEvents
                    gui()
                except Exception as e:
                    self.PromptBox.setText('报错：'+str(e))
                    gui=QApplication.processEvents
                    gui()
            except Exception as e:
                self.PromptBox.setText('报错：'+str(e))
                gui=QApplication.processEvents
                gui()
        return
    #渲染函数
    def GenReport(self):
        '''
        记录统计
        '''
        ArknightsGachaReport.PromptBox.setText('……正在执行可视化渲染……')
        gui=QApplication.processEvents
        gui()
        #记录初始化
        for s_name in agent_search:
            agent_search[s_name]={3: 0, 4: 0, 5: 0, 6: 0}
            agent_search_chars[s_name]=[]
            agent_search_chars_5star[s_name]=[]
            agent_search_cnt[s_name]=1
            agent_search_cnt_5star[s_name]=1
        #数据统计
        for i in range(1,df.shape[0]+1):
            rarity=df.loc[i]['星级'] 
            if df.loc[i]['寻访类别'] not in agent_search:
                df.loc[i,'寻访类别']='标准寻访'
            s_name=df.loc[i]['寻访类别']
            #星级计数
            agent_search['全部寻访'][rarity]+=1 
            agent_search[s_name][rarity]+=1
            #记录六星干员
            if rarity==6:
                agent_search_chars[s_name].append(df.loc[i]['干员']+"[%d]"%agent_search_cnt[s_name])
                agent_search_chars['全部寻访'].append(df.loc[i]['干员']+"[%d]"%agent_search_cnt['全部寻访'])
            #记录五星干员
            elif rarity==5:
                agent_search_chars_5star[s_name].append(df.loc[i]['干员']+"[%d]"%agent_search_cnt_5star[s_name])
                agent_search_chars_5star['全部寻访'].append(df.loc[i]['干员']+"[%d]"%agent_search_cnt_5star['全部寻访'])
            #六星水位记录
            df.loc[i,'全部寻访累计未出六星次数']=agent_search_cnt['全部寻访']
            df.loc[i,'所属寻访类别中累计未出六星次数']=agent_search_cnt[s_name]
            agent_search_cnt['全部寻访']=agent_search_cnt['全部寻访']+1 if rarity!=6 else 1
            agent_search_cnt[s_name]=agent_search_cnt[s_name]+1 if rarity!=6 else 1
            #五星水位
            df.loc[i,'全部寻访累计未出五星次数']=agent_search_cnt_5star['全部寻访']
            df.loc[i,'所属寻访类别中累计未出五星次数']=agent_search_cnt_5star[s_name]
            agent_search_cnt_5star['全部寻访']=agent_search_cnt_5star['全部寻访']+1 if rarity<5 else 1
            agent_search_cnt_5star[s_name]=agent_search_cnt_5star[s_name]+1 if rarity<5 else 1
        '''
        可视化渲染
        ''' 
        #可视化渲染
        from pyecharts import options as opts
        from pyecharts.charts import Pie, Timeline
        from pyecharts.globals import ThemeType
        win_width,win_height=str(0.9*self.geometry().width())+'px',str(0.8*self.geometry().height())+'px'
        tl =Timeline(init_opts=opts.InitOpts(width=win_width, height=win_height,theme="macarons"))#,theme=ThemeType.MACARONS
        tl.add_schema(pos_left="3%", pos_right="3%",pos_bottom="7%",is_auto_play=True,play_interval=1000)
        for s_name in agent_search:
            star=agent_search[s_name]#星级字典
            try:
                avg=round((sum(agent_search[s_name].values())-agent_search_cnt[s_name]+1)/star[6],2)
            except:
                avg='?'
            try:
                avg_5=round((sum(agent_search[s_name].values())-agent_search_cnt_5star[s_name]+1)/star[5],2)
            except:
                avg_5='?'
            pie = (
                Pie(init_opts=opts.InitOpts(theme=ThemeType.MACARONS))
                .add(
                    s_name,
                    [list(z) for z in zip([str(key)+"星" for key in star.keys()],star.values())],
                    radius=["40%", "60%"],
                    center=["35%", "50%"],
                    label_opts=opts.LabelOpts(formatter="{b}: {c} \n ({d}%)")
                    )
                .set_global_opts(
                    title_opts=opts.TitleOpts(title=s_name+"["+str(sum(agent_search[s_name].values()))+"]",subtitle="距上次六星{}次,距上次五星{}次\n六星平均需{}抽，五星平均需{}抽".format(agent_search_cnt[s_name]-1,agent_search_cnt_5star[s_name]-1,avg,avg_5),pos_left="5%",pos_top="5%"),
                    #图例排布
                    legend_opts=opts.LegendOpts(type_="scroll", pos_left="75%",pos_bottom="center", orient="vertical"),
                    #文本框设置
                    graphic_opts=[
                            opts.GraphicGroup(
                                graphic_item=opts.GraphicItem(left="31%",top="center",), # 控制整体的位置
                                children=[
                                    # opts.GraphicText控制文字的显示
                                    opts.GraphicText(
                                        graphic_item=opts.GraphicItem(left="center",top="middle",z=100,is_silent=False,is_draggable=True),
                                        graphic_textstyle_opts=opts.GraphicTextStyleOpts(
                                            # 可以通过jsCode添加js代码，也可以直接用字符串
                                            text="六星历史记录\n"+"\n".join(agent_search_chars[s_name][-10:]) if len(agent_search_chars[s_name])>0 else "六星历史记录\n\n"+"暂无",
                                            font="{}px HGXBS_CNKI".format(str(16*self.geometry().width()/1080)),
                                            text_align='left',text_vertical_align='middle',
                                            graphic_basicstyle_opts=opts.GraphicBasicStyleOpts(fill= "#666")
                                        )
                                    )
                                ]
                            ),
                            opts.GraphicGroup(
                                graphic_item=opts.GraphicItem(left="64%",top="center",), # 控制整体的位置
                                children=[
                                    opts.GraphicText(
                                        graphic_item=opts.GraphicItem(left="center",top="middle",z=100,is_silent=False,is_draggable=True),
                                        graphic_textstyle_opts=opts.GraphicTextStyleOpts(
                                            text="五星历史记录\n"+"\n".join(agent_search_chars_5star[s_name][-20:]) if len(agent_search_chars_5star[s_name])>0 else "五星历史记录\n\n"+"暂无",
                                            font="{}px HGSS_CNKI".format(str(14*self.geometry().width()/1080)),
                                            text_align='left',text_vertical_align='middle',
                                            graphic_basicstyle_opts=opts.GraphicBasicStyleOpts(fill= "#666")
                                            )
                                    )
                                ]
                            )
                    ]
                )
                .set_series_opts(color=["#5AB1EF","#B6A2DE","#FFB980","#F55066"],tooltip_opts=opts.TooltipOpts(trigger="item", formatter="{a} <br/>{b}: {c} ({d}%)")
                )
            )
            tl.add(chart=pie,time_point=s_name)
        tl.render("./UserData/ArknightsGachaReport.html")
        '''
        界面更新
        '''
        #为防止产生进程冲突，先尝试关闭excel.exe
        wmi = GetObject('winmgmts:')
        processCodeCov = wmi.ExecQuery('select * from Win32_Process where name=\"%s\"' %('EXCEL.EXE'))
        if len(processCodeCov) > 0:
            os.system('%s%s' % ("taskkill /F /IM ",'EXCEL.EXE'))
        #确认不会产生进程冲突，存储excel文件
        df.to_excel(file_name,index = False)
        self.webBrowser.load(QUrl(QFileInfo("./UserData/ArknightsGachaReport.html").absoluteFilePath()))
        self.PromptBox.setText( '渲染完毕')
        gui=QApplication.processEvents
        gui()
        return

class GachaDataView(QWidget, Ui_Form):
    def __init__(self, parent=None):    
        super(GachaDataView, self).__init__(parent)
        self.setupUi(self)
        self.retranslateUi(self)
        self.setWindowIcon(QIcon(':Ark.ico'))#图标设置
        self.setWindowTitle("明日方舟寻访数据")
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidget.setSelectionBehavior (QAbstractItemView. SelectRows)
        self.child1=NoteSearchWindow()#子窗口1
        self.child2=RemoveNoteWindow()#子窗口2
        try:
            self.Renew(df)
        except NameError:
            ArknightsGachaReport.PromptBox.setText('……尚未取得数据，请点击[开始分析]取得数据……')
            gui=QApplication.processEvents
            gui()
        self.center()
    #窗口居中
    def center(self):  
        screen = QDesktopWidget().screenGeometry()  
        size = self.geometry()    
        self.move((screen.width() - size.width()) / 2,  (screen.height() - size.height()) / 2)
    #根据df数据更新tableWidget视图 
    def Renew(self,df):  
        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1]-2)
        self.tableWidget.setHorizontalHeaderLabels(df.columns)
        for i in range(df.shape[0]):
            for j in range(df.shape[1]-2):
                newItem = QTableWidgetItem(str(df.iloc[i][j]))  
                self.tableWidget.setItem(i, j, newItem)
        self.tableWidget.resizeColumnsToContents ()
        self.tableWidget.resizeRowsToContents ()
        return
    #寻访标注
    def NoteSearch(self):
        if self.WorkLayout.indexOf(self.child1)==-1:
            self.WorkLayout.addWidget(self.child1)
        self.child1.show ()
        return
    #标注移除
    def RemoveNote(self):
        if self.WorkLayout.indexOf(self.child2)==-1:
            self.WorkLayout.addWidget(self.child2)
        self.child2.show ()
        return

class NoteSearchWindow(QWidget, Ui_NoteSearchWindow):
    def __init__(self, parent=None):    
        super(NoteSearchWindow, self).__init__(parent)
        self.setupUi(self)
        self.retranslateUi(self)
        intValidator=QIntValidator(self)
        intValidator.setRange(1,df.shape[0])
        self.BeginEdit.setValidator(intValidator)
        self.EndEdit.setValidator(intValidator)
    def SendNote(self):
        s_name=self.NameEdit.text()
        if s_name not in agent_search:#如果寻访名称没有被记录就进行记录
            agent_search[s_name]={3:0,4:0,5:0,6:0}
        begin,end=int(self.BeginEdit.text()),int(self.EndEdit.text())
        index=set(df.index)
        for i in range(begin,end+1):
            if i in index:
                df.loc[i,'寻访类别']=s_name
        ArknightsGachaReport.DataWindow.Renew(df)
        ArknightsGachaReport.GenReport()
        return
class RemoveNoteWindow(QWidget, Ui_RemoveNote):
    def __init__(self, parent=None):    
        super(RemoveNoteWindow, self).__init__(parent)
        self.setupUi(self)
        self.retranslateUi(self)
        self.comboBox.addItems([key for key in agent_search.keys() if key not in ['全部寻访','标准寻访']])
    def SendRemove(self):
        s_name=self.comboBox.currentText()
        try:
            del agent_search[s_name]
        except:
            pass
        ArknightsGachaReport.DataWindow.Renew(df)
        ArknightsGachaReport.GenReport()
        return
#执行main函数
if __name__=="__main__":  
    app = QApplication(sys.argv)  
    ArknightsGachaReport = ArknightsGachaReportWindow()  
    ArknightsGachaReport.show()
    sys.exit(app.exec_())  