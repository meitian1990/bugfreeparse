import os,xlrd,xlwt
from xlutils.copy import copy
class parsebugfree():
    #下面的num为统计的一些指标数据
    smokenum=0#冒烟bug数
    interfacenum=0#接口bug数
    functionnum=0#功能bug数
    servernum=0#server端bug数
    iosnum=0#ios端bug数
    Androidnum=0#Android端bug数
    FEnum=0#FE端bug数
    othernum=0#未知开发的bug数，需要补全user.txt
    requirementschangenum=0#需求变更的bug数
    legacynum=0#遗留的bug数
    notnugnum=0#不是bug的数量
    unablereproducebugnum=0#无法重现bug数量
    repeatbugnum=0#重复bug数量
    specialnum=0#专项bug数量
    
    #每日创建bug数
    createddate=dict()#存储每日的bug创建数量
    ioscreateddate=dict()
    Androidcreateddate=dict()
    servercreateddate=dict()
    FEcreateddate=dict()
    unknowncreateddate=dict()
    #每日解决bug数
    reslovedate=dict()#存储每日的bug解决数量
    iosreslovedate=dict()
    Androidreslovedate=dict()
    serverreslovedate=dict()
    FEreslovedate=dict()
    unknownreslovedate=dict()
    #每个人每天创建的bug数量
    ios_everybody_bugnum_created=dict()#ios每个人每天创建的bug数量
    Android_everybody_bugnum_created=dict()#Android每个人每天创建的bug数量
    server_everybody_bugnum_created=dict()#server每个人每天创建的bug数量
    FE_everybody_bugnum_created=dict()#FE每个人每天创建的bug数量
    #每个人每天解决的bug数量
    ios_everybody_bugnum_reslove=dict()#ios每个人每天解决的bug数量
    Android_everybody_bugnum_reslove=dict()#Android每个人每天解决的bug数量
    server_everybody_bugnum_reslove=dict()#server每个人每天解决的bug数量
    FE_everybody_bugnum_reslove=dict()#FE每个人每天解决的bug数量
    #下面的数据为统计BUG.xml的每个sheet页的行数，方便插入数据 
    sheet1num=1
    sheet2num=1
    sheet3num=1
    sheet4num=1
    sheet5num=1
    sheet6num=1
    sheet7num=1
    sheet8num=1
    sheet9num=1
    sheet10num=1
    sheet15num=1
    
    #获得当前路径下所有的excel文件，实际上为要处理的到处的bug的文档
    def getexcelfiles(self):
        files=[x for x in os.listdir("./") if ".xls" in x ]
        print("当前目录下有以下要处理的bug表：\n",files)
        return files

    #获得开发名单
    def getuser(self):
        user=dict()
        with open ("./user.txt") as f:          
            for line in f:
                _=line.strip().split("，")
                user[_[0]]=_[1]
        #print(user)
        return user
    def cellstype(self):#设置标题的样式
        style = xlwt.XFStyle()
        style.pattern.pattern = 1
        style.pattern.pattern_fore_colour = 3
        style.borders.bottom = 1
        g_headerFont = xlwt.Font()
        g_headerFont.bold = True
        style.font = g_headerFont
        return style
    #定义sheet的列宽，sheet1表示sheet，n表示多少列
    def celwidth(self,sheet1,n,width=256*20):
        for i in range(n):
            sheet1.col(i).width=width

    #重写excel数据，根据解决方案分表
    def newrowdata(self,sheet,row,data0,data1,data2,data3,data4,data5,data6,data7,data8,data9):#为sheet插入9个数据，因为分表后一共就保存9个数据，所以这么写。
        sheet.write(row,0,data0)
        sheet.write(row,1,data1)
        sheet.write(row,2,data2)
        sheet.write(row,3,data3)
        sheet.write(row,4,data4)
        sheet.write(row,5,data5)
        sheet.write(row,6,data6)
        sheet.write(row,7,data7)
        sheet.write(row,8,data8)
        sheet.write(row,9,data9)
    def newsheet(self,workbook,name):#为每个sheet页第一行建一个title，所有sheet页是一样的
        sheet=workbook.add_sheet(name,cell_overwrite_ok=True)
        sheet.write(0,0,"需求名称",self.cellstype())
        sheet.write(0,1,"需求ID",self.cellstype())
        sheet.write(0,2,"BUGID",self.cellstype())
        sheet.write(0,3,"主题",self.cellstype())
        sheet.write(0,4,"经办人",self.cellstype())
        sheet.write(0,5,"问题解决人",self.cellstype())
        sheet.write(0,6,"解决时间",self.cellstype())
        sheet.write(0,7,"创建时间",self.cellstype())
        sheet.write(0,8,"BUG解决方案",self.cellstype())
        sheet.write(0,9,"BUG状态",self.cellstype())
        self.celwidth(sheet,10)
        return sheet
    def overwriteexcel(self,files):
        f=xlwt.Workbook()
        sheet1=f.add_sheet("BUG统计",cell_overwrite_ok=True)
        sheet2=self.newsheet(f,"有效BUG")
        sheet3=self.newsheet(f,"未知解决人BUG，需手动处理")
        sheet4=self.newsheet(f,"冒烟测试")
        sheet5=self.newsheet(f,"接口测试")
        sheet6=self.newsheet(f,"需求变更导致BUG")
        sheet7=self.newsheet(f,"遗留BUG")
        sheet8=self.newsheet(f,"不是BUG")
        sheet9=self.newsheet(f,"重复BUG")
        sheet10=self.newsheet(f,"无法重现BUG")
        sheet11=f.add_sheet("ios端每个人的BUG数",cell_overwrite_ok=True)
        sheet11.write(0,0,"开发姓名")
        sheet11.write(0,1,"BUG数量")
        self.celwidth(sheet11,2)
        sheet12=f.add_sheet("Android端每个人的BUG数",cell_overwrite_ok=True)
        sheet12.write(0,0,"开发姓名")
        sheet12.write(0,1,"BUG数量")
        self.celwidth(sheet12,2)
        sheet13=f.add_sheet("server端每个人的BUG数",cell_overwrite_ok=True)
        sheet13.write(0,0,"开发姓名")
        sheet13.write(0,1,"BUG数量")
        self.celwidth(sheet13,2)
        sheet14=f.add_sheet("FE端每个人的BUG数",cell_overwrite_ok=True)
        sheet14.write(0,0,"开发姓名")
        sheet14.write(0,1,"BUG数量")
        self.celwidth(sheet14,2)
        sheet15=self.newsheet(f,"重新打开的BUG数")
        sheet15.write(0,10,"BUG重新打开次数",self.cellstype())

        for i in files:
            workbook = xlrd.open_workbook(i)
            sheet=workbook.sheet_by_index(0)
            #print(sheet.name,sheet.nrows,sheet.ncols)
            print("正在处理的excel：",i)
            for i in range(2,sheet.nrows):
                try:
                    data=dict()
                    data["需求名称"]=sheet.cell(i,5).value
                    data["需求ID"]=sheet.cell(i,4).value
                    data["BUGID"]=sheet.cell(i,6).value
                    data["主题"]=sheet.cell(i,7).value
                    data["经办人"]=sheet.cell(i,11).value
                    data["问题解决人"]=sheet.cell(i,12).value
                    data["解决时间"]=sheet.cell(i,13).value[:10]
                    data["创建时间"]=sheet.cell(i,15).value[:10]
                    data["BUG解决方案"]=sheet.cell(i,18).value
                    data["BUG状态"]=sheet.cell(i,8).value
                    data["BUG重新打开次数"]=sheet.cell(i,10).value
                    #print(data,"\n")
                except Exception as e:
                    print(e)
                    print("数据有问题的行：",i)
                    break
                
                if data["BUG解决方案"]=="已解决" or data["BUG解决方案"]=="未解决":
                    self.newrowdata(sheet2,self.sheet2num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])
                    self.sheet2num=self.sheet2num+1
                    
                    if int(data["BUG重新打开次数"])>0:
                        self.newrowdata(sheet15,self.sheet15num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])
                        sheet15.write(self.sheet15num,10,data["BUG重新打开次数"])
                        self.sheet15num=self.sheet15num+1
                    
                    if "【冒烟】" in data["主题"] or "[冒烟]" in data["主题"] or "【冒烟未通过】" in data["主题"] or "[冒烟未通过]" in data["主题"]:
                        self.newrowdata(sheet4,self.sheet4num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet4num=self.sheet4num+1
                    elif "【接口】" in data["主题"] or "[接口]" in data["主题"] or "【接口测试】" in data["主题"] or "[接口测试]" in data["主题"]:
                        self.newrowdata(sheet5,self.sheet5num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet5num=self.sheet5num+1
                    elif "[专项测试]" in data["主题"] or "【专项测试】" in data["主题"]:
                        self.specialnum=self.specialnum+1

                elif data["BUG解决方案"]=="不是Bug":
                    self.newrowdata(sheet8,self.sheet8num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])                
                    self.sheet8num=self.sheet8num+1
                elif data["BUG解决方案"]=="重复":
                    self.newrowdata(sheet9,self.sheet9num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])                
                    self.sheet9num=self.sheet9num+1
                elif data["BUG解决方案"]=="无法重现":
                    self.newrowdata(sheet10,self.sheet10num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])                
                    self.sheet10num=self.sheet10num+1
                elif data["BUG解决方案"]=="需求变更导致":
                    self.newrowdata(sheet6,self.sheet6num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])                 
                    self.sheet6num=self.sheet6num+1
                    self.newrowdata(sheet2,self.sheet2num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])
                    self.sheet2num=self.sheet2num+1
                    
                    if int(data["BUG重新打开次数"])>0:
                        self.newrowdata(sheet15,self.sheet15num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])
                        sheet15.write(self.sheet15num,10,data["BUG重新打开次数"])
                        self.sheet15num=self.sheet15num+1
                    
                    if "【冒烟】" in data["主题"] or "[冒烟]" in data["主题"] or "【冒烟未通过】" in data["主题"] or "[冒烟未通过]" in data["主题"]:
                        self.newrowdata(sheet4,self.sheet4num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet4num=self.sheet4num+1
                    elif "【接口】" in data["主题"] or "[接口]" in data["主题"] or "【接口测试】" in data["主题"] or "[接口测试]" in data["主题"]:
                        self.newrowdata(sheet5,self.sheet5num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet5num=self.sheet5num+1
                    elif "[专项测试]" in data["主题"] or "【专项测试】" in data["主题"]:
                        self.specialnum=self.specialnum+1
                        
                elif data["BUG解决方案"]=="遗留" or data["BUG解决方案"]=="以后解决":
                    self.newrowdata(sheet7,self.sheet7num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])                
                    self.sheet7num=self.sheet7num+1
                    self.newrowdata(sheet2,self.sheet2num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])
                    self.sheet2num=self.sheet2num+1
                    
                    if int(data["BUG重新打开次数"])>0:
                        self.newrowdata(sheet15,self.sheet15num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])
                        sheet15.write(self.sheet15num,10,data["BUG重新打开次数"])
                        self.sheet15num=self.sheet15num+1
                    
                    if "【冒烟】" in data["主题"] or "[冒烟]" in data["主题"] or "【冒烟未通过】" in data["主题"] or "[冒烟未通过]" in data["主题"]:
                        self.newrowdata(sheet4,self.sheet4num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet4num=self.sheet4num+1
                    elif "【接口】" in data["主题"] or "[接口]" in data["主题"] or "【接口测试】" in data["主题"] or "[接口测试]" in data["主题"]:
                        self.newrowdata(sheet5,self.sheet5num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet5num=self.sheet5num+1
                    elif "[专项测试]" in data["主题"] or "【专项测试】" in data["主题"]:
                        self.specialnum=self.specialnum+1
                else:
                    print("未知的解决方案：",data["BUG解决方案"],"；对应的BUGID：",data["BUGID"])
                f.save("./BUG分析.xls")
        print("\n***************BUG分表结束，下面进行数据统计***************")
        
    #数据分析，对分表后的结果进行统计 
    def dataanalysis(self):
        user=self.getuser()#获得user.txt中的开发名单
        workbook = xlrd.open_workbook("./BUG分析.xls")
        sheet2=workbook.sheet_by_name("有效BUG")
        sheet4=workbook.sheet_by_name("冒烟测试")
        self.smokenum=sheet4.nrows-1
        sheet5=workbook.sheet_by_name("接口测试")
        self.interfacenum=sheet5.nrows-1
        self.functionnum=sheet2.nrows-1-self.interfacenum
        sheet6=workbook.sheet_by_name("需求变更导致BUG")
        self.requirementschangenum=sheet6.nrows-1
        sheet7=workbook.sheet_by_name("遗留BUG")
        self.legacynum=sheet7.nrows-1
        sheet8=workbook.sheet_by_name("不是BUG")
        self.notnugnum=sheet8.nrows-1#不是bug的数量
        sheet9=workbook.sheet_by_name("重复BUG")
        self.repeatbugnum=sheet9.nrows-1#重复bug数量     
        sheet10=workbook.sheet_by_name("无法重现BUG")
        self.unablereproducebugnum=sheet10.nrows-1#无法重现bug数量
        
        effectivebugnum=sheet2.nrows-1
    
        wb=copy(workbook)#Excel处理不支持直接更改已存在的文件，需要打开后copy一份，然后再同名保存，以后希望可以优化
        sheet1=wb.get_sheet(0)#获得BUG统计的sheet
        sheet3=wb.get_sheet(2)#获得未知解决人BUG，需手动处理的sheet
        sheet11=wb.get_sheet(10)#ios端每个人的BUG数
        sheet12=wb.get_sheet(11)#Android端每个人的BUG数
        sheet13=wb.get_sheet(12)#server端每个人的BUG数
        sheet14=wb.get_sheet(13)#FE端每个人的BUG数

        
        
        sheet1.write(0,0,"ios的BUG数",self.cellstype())
        sheet1.write(0,1,"Android的BUG数",self.cellstype())
        sheet1.write(0,2,"server的BUG数",self.cellstype())
        sheet1.write(0,3,"FE的BUG数",self.cellstype())
        sheet1.write(0,4,"冒烟BUG数",self.cellstype())
        sheet1.write(0,5,"接口的BUG数",self.cellstype())
        sheet1.write(0,6,"需求变更数",self.cellstype())
        sheet1.write(0,7,"遗留BUG数",self.cellstype())
        sheet1.write(0,8,"接口漏测率",self.cellstype())
        sheet1.write(0,9,"不是BUG数",self.cellstype())
        sheet1.write(0,10,"重复BUG数",self.cellstype())
        sheet1.write(0,11,"无法重现BUG数",self.cellstype())
        sheet1.write(0,12,"专项测试BUG数",self.cellstype())
        self.celwidth(sheet1,13)
##        for i in range(9):
##            sheet1.col(i) .width=256*20   

        print("\n【***************如果存在未知开发，建议补全当前目录的user.txt文件再进行数据统计***************】")
        
        for i in range(1,sheet2.nrows):
            data=dict()
            data["需求名称"]=sheet2.cell(i,0).value
            data["需求ID"]=sheet2.cell(i,1).value
            data["BUGID"]=sheet2.cell(i,2).value
            data["主题"]=sheet2.cell(i,3).value
            data["经办人"]=sheet2.cell(i,4).value
            data["问题解决人"]=sheet2.cell(i,5).value
            data["解决时间"]=sheet2.cell(i,6).value[:10]
            data["创建时间"]=sheet2.cell(i,7).value[:10]
            data["BUG解决方案"]=sheet2.cell(i,8).value
            data["BUG状态"]=sheet2.cell(i,9).value
            RD=data["问题解决人"]

            if RD=="":
                RD=data["经办人"]            
            try:
                if user[RD]=="ios":
                    self.iosnum=self.iosnum+1
                    try:
                        self.ios_everybody_bugnum_created[RD]=self.ios_everybody_bugnum_created[RD]+1
                    except KeyError:
                        self.ios_everybody_bugnum_created[RD]=1
                        
                elif user[RD]=="Android":
                    self.Androidnum=self.Androidnum+1
                    try:
                        self.Android_everybody_bugnum_created[RD]=self.Android_everybody_bugnum_created[RD]+1
                    except KeyError:
                        self.Android_everybody_bugnum_created[RD]=1
                        
                elif user[RD]=="FE":
                    self.FEnum=self.FEnum+1
                    try:
                        self.FE_everybody_bugnum_created[RD]=self.FE_everybody_bugnum_created[RD]+1
                    except KeyError:
                        self.FE_everybody_bugnum_created[RD]=1
                    
                elif user[RD]=="server":
                    self.servernum=self.servernum+1
                    try:
                        self.server_everybody_bugnum_created[RD]=self.server_everybody_bugnum_created[RD]+1
                    except KeyError:
                        self.server_everybody_bugnum_created[RD]=1
                        
            except KeyError:
                    print("未知的开发：",RD,"；对应的BUGID：",data["BUGID"])
                    self.othernum=self.othernum+1
                    self.newrowdata(sheet3,self.sheet3num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])                 
                    self.sheet3num=self.sheet3num+1
                    
            #按日统计每日bug解决数
            if data["解决时间"]!="":
                try:
                    self.reslovedate[data["解决时间"]]=self.reslovedate[data["解决时间"]]+1
                except KeyError:
                    print(data["解决时间"],"有BUG解决")
                    self.reslovedate[data["解决时间"]]=1
                try:
                    if user[RD]=="ios":
                        try:
                            self.iosreslovedate[data["解决时间"]]=self.iosreslovedate[data["解决时间"]]+1
                        except KeyError:
                            print(data["解决时间"],"ios有BUG解决") 
                            self.iosreslovedate[data["解决时间"]]=1
                    elif user[RD]=="Android":
                        try:
                            self.Androidreslovedate[data["解决时间"]]=self.Androidreslovedate[data["解决时间"]]+1
                        except KeyError:
                            print(data["解决时间"],"Android有BUG解决")
                            self.Androidreslovedate[data["解决时间"]]=1
                    elif user[RD]=="FE":
                        try:
                            self.FEreslovedate[data["解决时间"]]=self.FEreslovedate[data["解决时间"]]+1
                        except KeyError:
                            print(data["解决时间"],"FE有BUG解决")
                            self.FEreslovedate[data["解决时间"]]=1
                    elif user[RD]=="server":
                        try:
                            self.serverreslovedate[data["解决时间"]]=self.serverreslovedate[data["解决时间"]]+1
                        except KeyError:
                            print(data["解决时间"],"server有BUG解决")
                            self.serverreslovedate[data["解决时间"]]=1
                except KeyError:
                    print("未知端的bug解决人")
                    try:
                            self.unknownreslovedate[data["解决时间"]]=self.unknownreslovedate[data["解决时间"]]+1
                    except KeyError:
                            print(data["解决时间"],"未知端有BUG解决")
                            self.unknownreslovedate[data["解决时间"]]=1

##            _DEBUG = True
##            if _DEBUG == True: 
##                                import pdb 
##                                pdb.set_trace() 
            
            #按日统计每日创建bug数            
            try:
                self.createddate[data["创建时间"]]=self.createddate[data["创建时间"]]+1
            except KeyError:
                print(data["创建时间"],"有BUG产生")
                self.createddate[data["创建时间"]]=1
                
            try:
                if user[RD]=="ios":
                    try:
                        self.ioscreateddate[data["创建时间"]]=self.ioscreateddate[data["创建时间"]]+1
                    except KeyError:
                        print(data["创建时间"],"ios有BUG产生")
                        self.ioscreateddate[data["创建时间"]]=1
                elif user[RD]=="Android":
                    try:
                        self.Androidcreateddate[data["创建时间"]]=self.Androidcreateddate[data["创建时间"]]+1
                    except KeyError:
                        print(data["创建时间"],"Android有BUG产生")
                        self.Androidcreateddate[data["创建时间"]]=1
                elif user[RD]=="FE":
                    try:
                        self.FEcreateddate[data["创建时间"]]=self.FEcreateddate[data["创建时间"]]+1
                    except KeyError:
                        print(data["创建时间"],"FE有BUG产生")
                        self.FEcreateddate[data["创建时间"]]=1
                elif user[RD]=="server":
                    try:
                        self.servercreateddate[data["创建时间"]]=self.servercreateddate[data["创建时间"]]+1
                    except KeyError:
                        print(data["创建时间"],"server有BUG产生")
                        self.servercreateddate[data["创建时间"]]=1
            except KeyError:
                print("未知端的bug解决人")
                try:
                        self.unknowncreateddate[data["创建时间"]]=self.unknowncreateddate[data["创建时间"]]+1
                except KeyError:
                        print(data["创建时间"],"有BUG产生")
                        self.unknowncreateddate[data["创建时间"]]=1
            

        sheet1.write(1,0,self.iosnum)
        sheet1.write(1,1,self.Androidnum)
        sheet1.write(1,2,self.servernum)
        sheet1.write(1,3,self.FEnum)
        sheet1.write(1,4,self.smokenum)
        sheet1.write(1,5,self.interfacenum)
        sheet1.write(1,6,self.requirementschangenum)
        sheet1.write(1,7,self.legacynum)
        sheet1.write(1,8,(effectivebugnum-self.interfacenum)/effectivebugnum)
        sheet1.write(1,9,self.notnugnum)
        sheet1.write(1,10,self.repeatbugnum)
        sheet1.write(1,11,self.unablereproducebugnum)
        sheet1.write(1,12,self.specialnum)
    

        m=1#用于存储当前创建bug存储的列数
        sheet1.write(5,0,"日期")
        sheet1.write(6,0,"所有端统计")
        sheet1.write(7,0,"ios端统计")
        sheet1.write(8,0,"Android端统计")
        sheet1.write(9,0,"server端统计")
        sheet1.write(10,0,"FE端统计")
        sheet1.write(11,0,"未知端统计")
        
        keys_createddate=sorted(self.createddate.keys())
        for k in keys_createddate:
            sheet1.write(4,m,"",self.cellstype())
            sheet1.write(4,0,"按照BUG创建日期统计数量：",self.cellstype())
            sheet1.write(5,m,k)
            sheet1.write(6,m,self.createddate[k])
            try:
                sheet1.write(7,m,self.ioscreateddate[k])
            except:
                print("当前日期该端无数据！")
            try:
                sheet1.write(8,m,self.Androidcreateddate[k])
            except:
                print("当前日期该端无数据！")
            try:
                sheet1.write(9,m,self.servercreateddate[k])
            except:
                print("当前日期该端无数据！")
            try:
                sheet1.write(10,m,self.FEcreateddate[k])
            except:
                print("当前日期该端无数据！")
            try:
                sheet1.write(11,m,self.unknowncreateddate[k])
            except:
                print("当前日期该端无数据！")
            m=m+1
        

        n=1#用于存储当前解决bug日期存储的列数
        sheet1.write(14,0,"日期")
        sheet1.write(15,0,"所有端统计")
        sheet1.write(16,0,"ios端统计")
        sheet1.write(17,0,"Android端统计")
        sheet1.write(18,0,"server端统计")
        sheet1.write(19,0,"FE端统计")
        sheet1.write(20,0,"未知端统计")

        keys_reslovedate=sorted(self.reslovedate.keys())
        for k in keys_reslovedate:
            sheet1.write(13,n,"",self.cellstype())
            sheet1.write(13,0,"按照BUG解决日期统计数量：",self.cellstype())
            sheet1.write(14,n,k)
            sheet1.write(15,n,self.reslovedate[k])
            try:
                sheet1.write(16,n,self.iosreslovedate[k])
            except:
                print("当前日期该端无数据！")
            try:
                sheet1.write(17,n,self.Androidreslovedate[k])
            except:
                print("当前日期该端无数据！")
            try:
                sheet1.write(18,n,self.serverreslovedate[k])
            except:
                print("当前日期该端无数据！")
            try:
                sheet1.write(19,n,self.FEreslovedate[k])
            except:
                print("当前日期该端无数据！")
            try:
                sheet1.write(20,n,self.unknownreslovedate[k])
            except:
                print("当前日期该端无数据！")
            n=n+1
            
        ios_everybodybugnum=1#用于存储ios每个人的bug的行数
        for _ in self.ios_everybody_bugnum_created.keys():
            sheet11.write(ios_everybodybugnum,0,_)
            sheet11.write(ios_everybodybugnum,1,self.ios_everybody_bugnum_created[_])
            ios_everybodybugnum=ios_everybodybugnum+1

        Android_everybodybugnum=1#用于存储Android每个人的bug的行数
        for _ in self.Android_everybody_bugnum_created.keys():
            sheet12.write(Android_everybodybugnum,0,_)
            sheet12.write(Android_everybodybugnum,1,self.Android_everybody_bugnum_created[_])
            Android_everybodybugnum=Android_everybodybugnum+1

        server_everybodybugnum=1#用于存储server每个人的bug的行数
        for _ in self.server_everybody_bugnum_created.keys():
            sheet13.write(server_everybodybugnum,0,_)
            sheet13.write(server_everybodybugnum,1,self.server_everybody_bugnum_created[_])
            server_everybodybugnum=server_everybodybugnum+1

        FE_everybodybugnum=1#用于存储FE每个人的bug的行数
        for _ in self.FE_everybody_bugnum_created.keys():
            sheet14.write(FE_everybodybugnum,0,_)
            sheet14.write(FE_everybodybugnum,1,self.FE_everybody_bugnum_created[_])
            FE_everybodybugnum=FE_everybodybugnum+1

        
        
        #os.remove("./BUG分析.xls")
        wb.save("./BUG分析.xls")
        print("\n*****************数据分析完毕，打开当前目录的BUG分析.xls文件查看*****************")
        
if __name__=="__main__":
        if os.path.exists("./BUG分析.xls"):
            os.remove("./BUG分析.xls")
        my=parsebugfree()
        filelist=my.getexcelfiles()
        my.overwriteexcel(filelist)
        #user=my.getuser
        my.dataanalysis()
