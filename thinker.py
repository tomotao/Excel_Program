from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import os
import xlsxwriter
import pandas as pd
import xlrd
import time
import numpy as np

LOG_LINE_NUM=0
class MY_GUI():
    def __init__(self,init_window_name):
        self.init_window_name=init_window_name


    #设置窗口
    def set_init_window(self):
        self.init_window_name.title("Excel处理工具 by:三年")  # 设置窗口名
        self.init_window_name.geometry('1068x680+10+10')  # 1068x680为窗口大小,+10+10为窗口初始呈现位置
        self.init_window_name['bg'] = "white"  # 设置窗口背景颜色
        self.init_window_name.attributes("-alpha", 1.0)  # 设置窗口虚化程度


        self.top_frame = tk.Frame(self.init_window_name,bg='white') ##, height=20
        self.bottom_frame = tk.Frame(self.init_window_name,bg='white') ##, height=20
        self.top_sub1_frame = tk.Frame(self.top_frame,bg='white')
        self.top_sub2_frame = tk.Frame(self.top_frame)
        self.top_sub3_frame = tk.Frame(self.top_frame,bg='white')
        self.bottom_sub1_frame = tk.Frame(self.bottom_frame,bg='white')
        self.bottom_sub2_frame = tk.Frame(self.bottom_frame)
        self.bottom_sub3_frame = tk.Frame(self.bottom_frame,bg='white')


        self.top_frame.pack(fill='both')
        self.top_sub1_frame.pack(side='left',fill='y')
        self.top_sub2_frame.pack(side='left',fill='y')
        self.top_sub3_frame.pack(side='left',fill='y')
        self.bottom_frame.pack(fill='both')
        self.bottom_sub1_frame.pack(side='left',fill='y')
        self.bottom_sub2_frame.pack(side='left',fill='y')
        self.bottom_sub3_frame.pack(side='left',fill='y')

        # 设置标签
        self.init_data_label = Label(self.top_sub1_frame, text="待处理数据")
        self.log_label = Label(self.bottom_sub1_frame, text="事件日志")
        self.result_data_label = Label(self.top_sub3_frame, text="EnodeB结果预览")
        self.result_ecell_label = Label(self.bottom_sub3_frame, text="Ecell结果预览")

        self.init_data_label.pack(side='top',fill='x')
        self.log_label.pack(side='top',fill='x')
        self.result_data_label.pack(side='top',fill='x')
        self.result_ecell_label.pack(side='top',fill='x')

        # 设置文本框
        self.init_data_Text = Text(self.top_sub1_frame) # 原始数据录入框width=67, height=20
        self.result_data_Text = Text(self.top_sub3_frame) # 处理结果展示框width=67, height=20
        self.result_ecell_Text = Text(self.bottom_sub3_frame) # 处理结果展示框
        self.log_data_Text = Text(self.bottom_sub1_frame)# 日志框 width=67, height=30

        self.init_data_Text.pack(fill='both',expand='yes')
        self.result_data_Text.pack(fill='both',expand='yes')
        self.result_ecell_Text.pack(fill='both',expand='yes')
        self.log_data_Text.pack(fill='both',expand='yes')

        # 按钮
        self.str_trans_tomd4_button=Button(self.top_sub2_frame,text="Excel数据提取",bg='lightblue',width=10,command=self.main)#按钮调用内部方法，函数+()为直接调用
        self.str_trans_tomd5_button=Button(self.bottom_sub2_frame,text="生成结果",bg='lightblue',width=10,command=self.cell_data)#按钮调用内部方法，函数+()为直接调用
        self.str_trans_tomd4_button.pack(expand='yes')
        self.str_trans_tomd5_button.pack(expand='yes')

        # 加入tree代码
        self.tree = ttk.Treeview(self.result_data_Text, show="headings", selectmode='browse',height=14) ##show="tree"
        self.tree["columns"] = (
        "*ENODEB名称", "*ENODEBID", "*所属区县名称", "*所属站点名称", "*应急类型", "*蜂窝类型", "*VIP级别", "*软件版本", "*所属区县分公司名称", "*出厂时间",
        "*小微基站标识", "*设备类型")
        self.tree.column("*ENODEB名称", width=100)  # 表示列,不显示
        self.tree.column("*ENODEBID", width=100)
        self.tree.column("*所属区县名称", width=100)
        self.tree.column("*所属站点名称", width=100)  # 表示列,不显示
        self.tree.column("*应急类型", width=100)
        self.tree.column("*蜂窝类型", width=100)
        self.tree.column("*VIP级别", width=100)  # 表示列,不显示
        self.tree.column("*软件版本", width=100)
        self.tree.column("*所属区县分公司名称", width=100)
        self.tree.column("*出厂时间", width=100)  # 表示列,不显示
        self.tree.column("*小微基站标识", width=100)
        self.tree.column("*设备类型", width=100)
        self.tree.heading("*ENODEB名称", text="*ENODEB名称")  #
        self.tree.heading("*ENODEBID", text="*ENODEBID")
        self.tree.heading("*所属区县名称", text="*所属区县名称")
        self.tree.heading("*所属站点名称", text="*所属站点名称")  #
        self.tree.heading("*应急类型", text="*应急类型")
        self.tree.heading("*蜂窝类型", text="*蜂窝类型")
        self.tree.heading("*VIP级别", text="*VIP级别")  #
        self.tree.heading("*软件版本", text="*软件版本")
        self.tree.heading("*所属区县分公司名称", text="*所属区县分公司名称")
        self.tree.heading("*出厂时间", text="*出厂时间")  #
        self.tree.heading("*小微基站标识", text="*小微基站标识")
        self.tree.heading("*设备类型", text="*设备类型")




        self.vbary = ttk.Scrollbar(self.result_data_Text)
        self.vbarx = ttk.Scrollbar(self.result_data_Text, orient='horizontal')
        self.tree['yscrollcommand'] = self.vbary.set
        self.tree['xscrollcommand'] = self.vbarx.set

        self.vbary['command'] = self.tree.yview
        self.vbarx['command'] = self.tree.xview

        self.vbary.pack(side='right', fill='y')
        self.vbarx.pack(side='bottom', fill='x')
        self.tree.pack(fill='x',expand='yes')

        # 加入tree1代码
        self.tree1 = ttk.Treeview(self.result_ecell_Text, show="headings", selectmode='browse',height=100) ##show="tree"
        self.tree1["columns"] = (
        "*ECELL名称", "*小区码CI", "*跟踪区码TAC", "*所属ENODEB名称", "*覆盖场景", "*在用载频数量", "*行政区域类别", "*自动路测区域", "*优化片区责任人", "*经度",
        "*纬度", "*是否微基站")
        self.tree1.column("*ECELL名称", width=100)  # 表示列,不显示
        self.tree1.column("*小区码CI", width=100)
        self.tree1.column("*跟踪区码TAC", width=100)
        self.tree1.column("*所属ENODEB名称", width=100)  # 表示列,不显示
        self.tree1.column("*覆盖场景", width=100)
        self.tree1.column("*在用载频数量", width=100)
        self.tree1.column("*行政区域类别", width=100)  # 表示列,不显示
        self.tree1.column("*自动路测区域", width=100)
        self.tree1.column("*优化片区责任人", width=100)
        self.tree1.column("*经度", width=100)  # 表示列,不显示
        self.tree1.column("*纬度", width=100)
        self.tree1.column("*是否微基站", width=100)
        self.tree1.heading("*ECELL名称", text="*ECELL名称")  #
        self.tree1.heading("*小区码CI", text="*小区码CI")
        self.tree1.heading("*跟踪区码TAC", text="*跟踪区码TAC")
        self.tree1.heading("*所属ENODEB名称", text="*所属ENODEB名称")  #
        self.tree1.heading("*覆盖场景", text="*覆盖场景")
        self.tree1.heading("*在用载频数量", text="*在用载频数量")
        self.tree1.heading("*行政区域类别", text="*行政区域类别")  #
        self.tree1.heading("*自动路测区域", text="*自动路测区域")
        self.tree1.heading("*优化片区责任人", text="*优化片区责任人")
        self.tree1.heading("*经度", text="*经度")  #
        self.tree1.heading("*纬度", text="*纬度")
        self.tree1.heading("*是否微基站", text="*是否微基站")


        # tree1滚动条
        self.vbary1 = ttk.Scrollbar(self.result_ecell_Text)
        self.vbarx1 = ttk.Scrollbar(self.result_ecell_Text, orient='horizontal')
        self.tree1['yscrollcommand'] = self.vbary1.set
        self.tree1['xscrollcommand'] = self.vbarx1.set

        self.vbary1['command'] = self.tree1.yview
        self.vbarx1['command'] = self.tree1.xview

        self.vbary1.pack(side='right', fill='y')
        self.vbarx1.pack(side='bottom', fill='x')
        self.tree1.pack(fill='both',expand='yes')

    def main(self):
        global merge_annex4,merge_sheets,annex3_file,annex5_file,Plan_file,backfilling_data,new_path
        file_path = filedialog.askdirectory(title=u"请选择无线设备链接表路径")
        new_path = file_path + "/数据模板输出文件"
        backfilling_data = new_path + "/203-3、LTE链接.xlsx"
        self.makefile(new_path)
        merge_annex4 = new_path + "\附件4：合并区县数据.xlsx"
        annex4_file = filedialog.askopenfilename(title=u"请选择附件4的路径")
        merge_sheets = self.Excel_mergesheets(annex4_file)
        annex3_file = filedialog.askopenfilename(title=u"请选择附件3的路径")
        annex5_file = filedialog.askopenfilename(title=u"请选择附件5的路径")
        Plan_file = open(filedialog.askopenfilename(title=u"请选择统一规划.csv的路径"))
        self.write_data()
        self.write_log_to_Text("请稍等几分钟")
        self.init_data_Text.insert(END,"请在输入框内填入要输出的站点")
        # self.cell_data()
    def get_current_time(self):
        current_time=time.strftime('%Y-%m-%d  %H:%M:%S',time.localtime(time.time()))
        return current_time

    def write_log_to_Text(self,logmsg):
        global LOG_LINE_NUM
        current_time=self.get_current_time()
        logmsg_in=str(current_time)+" "+str(logmsg)+"\n" #换行
        if LOG_LINE_NUM<=7:
            self.log_data_Text.insert(END,logmsg_in)
            LOG_LINE_NUM=LOG_LINE_NUM+1
        else:
            self.log_data_Text.delete(1.0,2.0)
            self.log_data_Text.insert(END,logmsg_in)

    def makefile(self,path):
        if not os.path.exists(path):
            os.makedirs(path)
            workbook = xlsxwriter.Workbook(backfilling_data)
            worksheet_enodeb = workbook.add_worksheet("ENODEB")
            bold = workbook.add_format({'bold': 1})
            headings = ["*ENODEB名称", "OMC中网元名称", "*ENODEBID", "站号", "所属省份名称", "*所属地市名称", "*所属区县名称", "*所属站点名称",
                        "*所属机房名称",
                        "*X2 IP地址", "*S1 IP地址", "*所属OMC",
                        "所属MME名称", "所属SGW/SAEGE名称", "关联的MMEPOOL/SGSNPOOL", "关联的PGWPOOL", "关联的SAEGWPOOL/GGSNPOOL",
                        "关联的SGWPOOL", "*应急类型", "*蜂窝类型", "*VIP级别", "*工作频段",
                        "*厂商名称", "*规格型号", "*软件版本", "故障处理单位", "代维公司", "入网时间", "*所属区县分公司名称", "所属营业厅名称", "所属片区名称",
                        "工程名称",
                        "*出厂时间", "*小微基站标识",
                        "备注", "*是否城际站", "*构建方式", "*设备类型"]
            worksheet_enodeb.write_row('A1', headings, bold)
            worksheet_ecell = workbook.add_worksheet("ECELL")
            headings = ["*ECELL名称", "网管中网元名称", "*小区码CI", "*跟踪区码TAC", "所属省", "*所属地市", "所属区县", "所属站点名称", "所属机房名称",
                        "所属网络资源点名称", "*所属ENODEB名称", "边界小区类型",
                        "使用频段", "*覆盖类型", "*覆盖场景", "次要场景", "再次要场景", "二级场景", "室分覆盖区域", "*在用载频数量", "蜂窝类型", "入网时间",
                        "上下行子帧配比",
                        "特殊子帧时隙配比", "PCI", "*行政区域类别",
                        "*自动路测区域", "*优化片区责任人", "铁路沿线", "高速公路沿线", "*经度", "*纬度", "*是否微基站", ]
            worksheet_ecell.write_row('A1', headings, bold)
            workbook.close()
        self.write_log_to_Text("创建生成数据文件夹及输出excel模板完成！")

    #筛选基站的ip表格
    def Excel_mergesheets(self,excelFile):
        workbook = xlsxwriter.Workbook(merge_annex4)
        workbook.close()
        sum = []
        book = xlrd.open_workbook(excelFile)
        list_sheet = book.sheet_names()
        for j, sheetName in enumerate(list_sheet[2:16]):
            data = pd.read_excel(excelFile, sheetname=sheetName, encoding='utf-8')
            data_sheet = data[['基站名称', '基站IP地址1', '第2 IP']].dropna(axis=0, how='any')
            if j == 0:
                sum = data_sheet
            else:
                sum = pd.concat([sum, data_sheet], axis=0)
        sum.to_excel(merge_annex4, index=0)
        self.write_log_to_Text("附件4各区县数据合并完成，去选项框选择其他附件！")

    #完成Enodeb部分的数据计算
    def enodeb_data(self):
        annex3 = pd.read_excel(annex3_file)
        loc = annex3.loc[:, ['eNB ID（新10进制站号）', '站名(新', '归属区域', ')现场参考']]
        loc['*所属区县分公司名称'] = loc['归属区域'] + "分公司"
        loc['所属营业厅名称'] = loc['归属区域'] + "营业厅"
        print(loc.columns)
        loc.columns = ['*ENODEBID', '站号', '*所属区县名称', '*所属站点名称', '*所属区县分公司名称', "所属营业厅名称"]
        loc['*所属区县名称'] = loc['*所属区县名称'].replace(
            ['北仑', '慈溪', '海曙', '江北', '江东', '鄞州', '余姚', '镇海', '奉化', '宁海', '象山', '杭州'],
            ['北仑区', '慈溪市', '海曙区', '江北区', '鄞州区', '鄞州区', '余姚市', '镇海区', '奉化区', '宁海县', '象山县', '杭州市'])
        loc['所属省份名称'] = "浙江省"
        loc['*所属地市名称'] = "宁波市"
        loc['*所属OMC'] = "NB_LTE_NOK_OMCR1"
        loc['关联的MMEPOOL/SGSNPOOL'] = "宁舟-LTE-POOL1"
        loc['关联的SAEGWPOOL/GGSNPOOL'] = "诺基亚SAEGW"
        loc['*应急类型'] = "常规站"
        loc['*VIP级别'] = "非VIP"
        loc['*软件版本'] = "LNT3.1_ENB_1209_306_99"
        loc['所属片区名称'] = loc['*所属区县名称']
        loc['*ENODEB名称'] = loc['*所属站点名称']
        loc['*小微基站标识'] = np.where(["微RRU" in str(x) for x in loc["*所属站点名称"]], '是', '否')
        loc['*蜂窝类型'] = np.where([("SF" in str(x) or "W" in str(x)) for x in loc["*所属站点名称"]], '微蜂窝', '宏蜂窝')
        loc['入网时间'] = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        loc["*出厂时间"] = '%d-%02d-%02d' % (
            time.localtime().tm_year, time.localtime().tm_mon - 1, time.localtime().tm_mday)
        loc['*设备类型'] = np.where(["微RRU" in str(x) for x in loc["*所属站点名称"]], '微基站', '宏基站')
        backfilling = pd.read_excel(backfilling_data, sheetname="ENODEB")
        merge_annex4_data = pd.read_excel(merge_annex4)
        merge_annex4_data.columns = ['*所属站点名称', '*X2 IP地址', '*S1 IP地址']
        result_Temp = pd.merge(loc, merge_annex4_data, on='*所属站点名称', how='left')
        self.write_log_to_Text("完成sheet_name为ENODEB部分数据计算")
        print("完成sheet_name为ENODEB部分数据计算")
        return (result_Temp)

    #完成Ecell部分的数据计算
    def ecell_data(self):
        list = ["海曙区", "鄞州区", "江北区", "镇海区", "宁海县", "象山县", "奉化区", "余姚市", "北仑区", "慈溪市"]
        list_people = ["朱海明", "朱军", "钮泽敏", "徐雷", "林旭东", "杨韬勇", "张少庆", "林挺", "赖立人", "张雪锋"]
        annex5 = pd.read_excel(annex5_file, sheetname="数据申请站点")
        annex5_data = annex5[["ECI", "PCI", "TAC", '基站名称', '小区名', 'eNB']]
        annex5_data.columns = ['*小区码CI', 'PCI', '*跟踪区码TAC', '*所属ENODEB名称', '*ECELL名称', '所属站点名称']
        annex5_data["所属省"] = "浙江省"
        annex5_data["*所属地市"] = "宁波市"
        annex5_data["入网时间"] = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        annex5_data["蜂窝类型"] = np.where([("SF" in x or "W" in x) for x in annex5_data["*所属ENODEB名称"]], '微蜂窝',
                                       '宏蜂窝')
        annex5_data["*行政区域类别"] = np.where(
            [("镇海" in x or "宁海" in x or "象山" in x or "奉化" in x or "北仑" in x) for x in
             annex5_data["*所属ENODEB名称"]], '县城区', '市辖区')
        annex5_data['*是否微基站'] = np.where([("微RRU" in x) for x in annex5_data["*所属ENODEB名称"]], '是', '否')
        annex5_data['所属区县'] = [list[["海曙" in x, "鄞州" in x, "江北" in x, "镇海" in x, "宁海" in x, "象山" in x,
                                     "奉化" in x, "余姚" in x, "北仑" in x, "慈溪" in x].index(True)] for x in
                               annex5_data["*所属ENODEB名称"]]
        annex5_data['*优化片区责任人'] = [list_people[
                                       ["海曙" in x, "鄞州" in x, "江北" in x, "镇海" in x, "宁海" in x, "象山" in x,
                                        "奉化" in x, "余姚" in x, "北仑" in x, "慈溪" in x].index(True)] for x in
                                   annex5_data["*所属ENODEB名称"]]
        Plan_data = pd.read_csv(Plan_file, encoding='utf8')
        Plan_data_use = Plan_data.loc[:, ['NODEBID', 'BBU经度', 'BBU纬度', '覆盖场景']]
        Plan_data_use.columns = ["所属站点名称", "*经度", "*纬度", "*覆盖场景"]
        annex5_data = pd.merge(annex5_data, Plan_data_use, on='所属站点名称', how='left')
        annex5_data["所属站点名称"] = annex5_data["*所属ENODEB名称"]
        annex5_data["*在用载频数量"] = "1"
        annex5_data["上下行子帧配比"] = "1:3"
        annex5_data["特殊子帧时隙配比"] = "3-9-2"
        annex5_data["*自动路测区域"] = "18"
        self.write_log_to_Text("完成sheet为ECELL部分数据计算")
        print("完成sheet为ECELL部分数据计算")
        return (annex5_data)

    #把enodeB数据和Ecell数据进行合并并输出至203-3、LTE链接.xlsx表中
    def write_data(self):
        writer = pd.ExcelWriter(backfilling_data)
        backfilling = pd.read_excel(backfilling_data, sheetname="ENODEB")
        backfilling_sheet2 = pd.read_excel(backfilling_data, sheetname="ECELL")
        self.enodeb_data().to_excel(writer, sheet_name='ENODEB', columns=backfilling.columns, index=False)
        self.ecell_data().to_excel(writer, sheet_name="ECELL", columns=backfilling_sheet2.columns, index=False)
        writer.save()
        self.write_log_to_Text("完成数据写入203-3、LTE链接.xlsx")

    def cell_data(self):
        backfilling_data1 = filedialog.askopenfilename(title=u"请选择203-3、LTE链接.xlsx的路径")
        result_enodeb = pd.read_excel(backfilling_data1, sheetname="ENODEB", index=0)
        result_ecell = pd.read_excel(backfilling_data1, sheetname="ECELL", index=0)
        cell_select = self.init_data_Text.get(1.0,END).strip().replace("\n", "")
        cell_select=str(cell_select)
        print(cell_select)
        flag = 0
        while flag == 0:
            result_enodeb_cell = result_enodeb.loc[(result_enodeb['*所属站点名称']==cell_select)]
            # print(result_enodeb_cell)
            result_ecell_cell = result_ecell.loc[(result_ecell['*所属ENODEB名称'] == cell_select)]
            if result_enodeb_cell.empty and result_ecell_cell.empty:
                self.init_data_Text.insert(END, "输入的站点名称不存在或者输入内容有误，请重新输入！")
                cell_select = self.init_data_Text.get(END).strip().replace("\n", "")
                flag = 0
            else:
                flag = 1
        new_path1=filedialog.askdirectory(title=u"请选择输出数据指定文件夹路径")
        backfilling_file = new_path1 + "/" + cell_select + ".xlsx"
        writer = pd.ExcelWriter(backfilling_file)
        result_enodeb_cell.to_excel(writer, sheet_name='ENODEB', columns=result_enodeb.columns, index=False)
        result_ecell_cell.to_excel(writer, sheet_name="ECELL", columns=result_ecell.columns, index=False)
        writer.save()
        self.write_log_to_Text(cell_select + "内容输出完成")
        print(result_enodeb_cell.iloc[:,:],result_ecell_cell.iloc[:,:])
        try:
            self.tree1.insert("", 0, text="line1", values=(result_ecell_cell.iloc[1,0], result_ecell_cell.iloc[1,2], result_ecell_cell.iloc[1,3], result_ecell_cell.iloc[1,7],result_ecell_cell.iloc[1,14],result_ecell_cell.iloc[1,19],result_ecell_cell.iloc[1,25],result_ecell_cell.iloc[1,26],result_ecell_cell.iloc[1,27],result_ecell_cell.iloc[1,30],result_ecell_cell.iloc[1,31],result_ecell_cell.iloc[1,32]))
            self.tree1.insert("", 0, text="line1", values=(result_ecell_cell.iloc[2, 0], result_ecell_cell.iloc[2, 2], result_ecell_cell.iloc[2, 3],result_ecell_cell.iloc[2, 7], result_ecell_cell.iloc[2, 14], result_ecell_cell.iloc[2, 19],result_ecell_cell.iloc[2, 25], result_ecell_cell.iloc[2, 26], result_ecell_cell.iloc[2, 27], result_ecell_cell.iloc[2, 30], result_ecell_cell.iloc[2, 31], result_ecell_cell.iloc[2, 32]))
            self.tree1.insert("", 0, text="line1", values=(result_ecell_cell.iloc[0, 0], result_ecell_cell.iloc[0, 2], result_ecell_cell.iloc[0, 3],result_ecell_cell.iloc[0, 7], result_ecell_cell.iloc[0, 14], result_ecell_cell.iloc[0, 19],result_ecell_cell.iloc[0, 25], result_ecell_cell.iloc[0, 26], result_ecell_cell.iloc[0, 27],result_ecell_cell.iloc[0, 30], result_ecell_cell.iloc[0, 31], result_ecell_cell.iloc[0,32]))
            self.tree1.insert("", 0, text="line1", values=(result_ecell_cell.iloc[3, 0], result_ecell_cell.iloc[3, 2], result_ecell_cell.iloc[3, 3],result_ecell_cell.iloc[3, 7], result_ecell_cell.iloc[3, 14], result_ecell_cell.iloc[3, 19],result_ecell_cell.iloc[3, 25], result_ecell_cell.iloc[3, 26], result_ecell_cell.iloc[3, 27],result_ecell_cell.iloc[3, 30], result_ecell_cell.iloc[3, 31], result_ecell_cell.iloc[3, 32]))
            self.tree1.insert("", 0, text="line1", values=(result_ecell_cell.iloc[4, 0], result_ecell_cell.iloc[4, 2], result_ecell_cell.iloc[4, 3], result_ecell_cell.iloc[4, 7], result_ecell_cell.iloc[4, 14], result_ecell_cell.iloc[4, 19], result_ecell_cell.iloc[4, 25], result_ecell_cell.iloc[4, 26], result_ecell_cell.iloc[4, 27], result_ecell_cell.iloc[4, 30], result_ecell_cell.iloc[4, 31], result_ecell_cell.iloc[4, 32]))
            self.tree1.insert("", 0, text="line1", values=(result_ecell_cell.iloc[5, 0], result_ecell_cell.iloc[5, 2], result_ecell_cell.iloc[5, 3], result_ecell_cell.iloc[5, 7], result_ecell_cell.iloc[5, 14], result_ecell_cell.iloc[5, 19],result_ecell_cell.iloc[5, 25], result_ecell_cell.iloc[5, 26], result_ecell_cell.iloc[5, 27],result_ecell_cell.iloc[5, 30], result_ecell_cell.iloc[5, 31], result_ecell_cell.iloc[5, 32]))

        except:
            pass
        try:
            self.tree.insert("", 0, text="line1", values=(result_enodeb_cell.iloc[0,0],result_enodeb_cell.iloc[0,2],result_enodeb_cell.iloc[0,6],result_enodeb_cell.iloc[0,7],result_enodeb_cell.iloc[0,17],result_enodeb_cell.iloc[0,18],result_enodeb_cell.iloc[0,19],result_enodeb_cell.iloc[0,23],result_enodeb_cell.iloc[0,27],result_enodeb_cell.iloc[0,31],result_enodeb_cell.iloc[0,32],result_enodeb_cell.iloc[0,36],))  # 插入数据，
        except:
            pass

def gui_start( ):
    init_window=Tk()#实例化一个父窗口
    ZMJ_PORTAL=MY_GUI(init_window)
    ZMJ_PORTAL.set_init_window()

    init_window.mainloop()#父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示


if __name__ == '__main__':
    gui_start()