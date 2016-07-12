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
    createddate=dict()#存储每日的bug创建数量
    reslovedate=dict()#存储每日的bug解决数量
    #下面的数据为统计BUG.xml的每个sheet页的行数，方便插入数据
    sheet1num=1
    sheet2num=1
    sheet3num=1
    sheet4num=1
    sheet5num=1
    sheet6num=1
    sheet7num=1
    sheet8num=1
    
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
        sheet8=self.newsheet(f,"不是BUG，重复或无法重现BUG")
##        self.sheet1num=1
##        self.sheet2num=1
##        self.sheet3num=1
##        self.sheet4num=1
##        self.sheet5num=1
##        self.sheet6num=1
##        self.sheet7num=1
##        self.sheet8num=1
        for i in files:
            workbook = xlrd.open_workbook(i)
            sheet=workbook.sheet_by_index(0)
            #print(sheet.name,sheet.nrows,sheet.ncols)

            for i in range(2,sheet.nrows):
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
                #print(data,"\n")               
                if data["BUG解决方案"]=="已解决" or data["BUG解决方案"]=="未解决":
                    self.newrowdata(sheet2,self.sheet2num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])
                    self.sheet2num=self.sheet2num+1
                    if "冒烟" in data["主题"]:
                        self.newrowdata(sheet4,self.sheet4num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet4num=self.sheet4num+1
                    elif "接口" in data["主题"]:
                        self.newrowdata(sheet5,self.sheet2num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet5num=self.sheet5num+1

                elif data["BUG解决方案"]=="不是Bug" or data["BUG解决方案"]=="重复" or data["BUG解决方案"]=="无法重现":
                    self.newrowdata(sheet8,self.sheet8num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])                
                    self.sheet8num=self.sheet8num+1
                elif data["BUG解决方案"]=="需求变更导致":
                    self.newrowdata(sheet6,self.sheet6num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])                 
                    self.sheet6num=self.sheet6num+1
                    self.newrowdata(sheet2,self.sheet2num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])
                    self.sheet2num=self.sheet2num+1
                    if "冒烟" in data["主题"]:
                        self.newrowdata(sheet4,self.sheet4num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet4num=self.sheet4num+1
                    elif "接口" in data["主题"]:
                        self.newrowdata(sheet5,self.sheet5num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet5num=self.sheet5num+1
                elif data["BUG解决方案"]=="遗留" or data["BUG解决方案"]=="以后解决":
                    self.newrowdata(sheet7,self.sheet7num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])                
                    self.sheet7num=self.sheet7num+1
                    self.newrowdata(sheet2,self.sheet2num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])
                    self.sheet2num=self.sheet2num+1
                    if "冒烟" in data["主题"]:
                        self.newrowdata(sheet4,self.sheet4num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet4num=self.sheet4num+1
                    elif "接口" in data["主题"]:
                        self.newrowdata(sheet5,self.sheet5num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])   
                        self.sheet5num=self.sheet5num+1
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

        effectivebugnum=sheet2.nrows-1
    
        wb=copy(workbook)#Excel处理不支持直接更改已存在的文件，需要打开后copy一份，然后再同名保存，以后希望可以优化
        sheet1=wb.get_sheet(0)#获得BUG统计的sheet
        sheet3=wb.get_sheet(2)#获得未知解决人BUG，需手动处理的sheet
        dir(sheet1)
        sheet1.write(0,0,"ios的BUG数",self.cellstype())
        sheet1.write(0,1,"Android的BUG数",self.cellstype())
        sheet1.write(0,2,"server的BUG数",self.cellstype())
        sheet1.write(0,3,"FE的BUG数",self.cellstype())
        sheet1.write(0,4,"冒烟BUG数",self.cellstype())
        sheet1.write(0,5,"接口的BUG数",self.cellstype())
        sheet1.write(0,6,"需求变更数",self.cellstype())
        sheet1.write(0,7,"遗留BUG数",self.cellstype())
        sheet1.write(0,8,"接口漏测率",self.cellstype())
        for i in range(9):
            sheet1.col(i) .width=256*20   

        print("\n【***************如果存在未知开发，建议补全当前目录的user.txt文件再进行数据统计***************】")
        
        for i in range(2,sheet2.nrows):
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
                elif user[RD]=="Android":
                    self.Androidnum=self.Androidnum+1
                elif user[RD]=="FE":
                    self.FEnum=self.FEnum+1
                elif user[RD]=="server":
                    self.servernum=self.servernum+1
            except KeyError:
                    print("未知的开发：",RD,"；对应的BUGID：",data["BUGID"])
                    self.othernum=self.othernum+1
                    self.newrowdata(sheet3,self.sheet3num,data["需求名称"],data["需求ID"],data["BUGID"],data["主题"],data["经办人"],data["问题解决人"],data["解决时间"],data["创建时间"],data["BUG解决方案"],data["BUG状态"])                 
                    self.sheet3num=self.sheet3num+1
            if data["解决时间"]!="":
                try:
                    self.reslovedate[data["解决时间"]]=self.reslovedate[data["解决时间"]]+1
                except KeyError:
                    print(data["解决时间"],"有BUG产生")
                    self.reslovedate[data["解决时间"]]=1
            try:
                self.createddate[data["创建时间"]]=self.createddate[data["创建时间"]]+1
            except KeyError:
                print(data["创建时间"],"有BUG产生")
                self.createddate[data["创建时间"]]=1

        sheet1.write(1,0,self.iosnum)
        sheet1.write(1,1,self.Androidnum)
        sheet1.write(1,2,self.servernum)
        sheet1.write(1,3,self.FEnum)
        sheet1.write(1,4,self.smokenum)
        sheet1.write(1,5,self.interfacenum)
        sheet1.write(1,6,self.requirementschangenum)
        sheet1.write(1,7,self.legacynum)
        sheet1.write(1,8,self.interfacenum/effectivebugnum)

        m=0#用于存储当创建bug存储的列数
        for k in self.createddate.keys():
            sheet1.write(4,m,"",self.cellstype())
            sheet1.write(4,0,"按照BUG创建日期统计数量：",self.cellstype())
            sheet1.write(5,m,k)
            sheet1.write(6,m,self.createddate[k])
            m=m+1

        n=0#用于存储当解决bug日期存储的列数
        for k in self.reslovedate.keys():
            sheet1.write(8,n,"",self.cellstype())
            sheet1.write(8,0,"按照BUG解决日期统计数量：",self.cellstype())
            sheet1.write(9,n,k)
            sheet1.write(10,n,self.reslovedate[k])
            n=n+1

        
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
