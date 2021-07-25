from selenium import webdriver
import time,os,xlwt,sys,subprocess

tall_style = xlwt.easyxf('font:height 360')  # 36p

class splider:
    def __init__(self):
        self.url="http://data.eastmoney.com/zjlx/detail.html"
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.ShenA, self.HuA, self.ChuangA=[],[],[]
    #主函数
    def major(self):
        self.driver.implicitly_wait(10)
        self.driver.get(self.url)
        #点击深A板块
        js='''document.querySelector("#filter_mkt > li:nth-child(5)").click()'''
        self.driver.execute_script(js)
        time.sleep(6)
        self.ShenA=self.getAbankuai()

        #点击沪A板块
        js='''document.querySelector("#filter_mkt > li:nth-child(3)").click()'''
        self.driver.execute_script(js)
        js='''document.querySelector('div [data-value="sha"]').click()'''
        self.driver.execute_script(js)
        time.sleep(6)
        self.HuA=self.getAbankuai()

        # 点击创A板块
        js='''document.querySelector("#filter_mkt > li:nth-child(6)").click()'''
        self.driver.execute_script(js)
        js='''document.querySelector('div [data-value="cyb"]').click()'''
        self.driver.execute_script(js)
        time.sleep(6)
        self.ChuangA = self.getAbankuai()

        #转变数据
        self.ShenA=self.url2data(self.ShenA)
        self.HuA=self.url2data(self.HuA)
        self.ChuangA=self.url2data(self.ChuangA)
        #
        # self.ShenA=[[['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15'],['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15']],[['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15'],['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15']]]
        # self.HuA=[[['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15'],['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15']],[['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15'],['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15']]]
        # self.ChuangA=[[['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15'],['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15']],[['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15'],['300612', '宣亚国际', '-3.33%', '商务服务业', '2017-02-15']]]

        self.ShenA=self.data2writdata(self.ShenA)
        self.HuA=self.data2writdata(self.HuA)
        self.ChuangA=self.data2writdata(self.ChuangA)

        ctime=time.strftime("%Y-%m-%d %H%M", time.localtime())
        path="市场日报"+ctime+".xls"
        nameList=['沪A','深A','创A']

        wvalueList=[self.HuA,self.ShenA,self.ChuangA]
        self.write_excel_xls(path,nameList,wvalueList)

    #url转化成值
    def url2data(self,listdata):
        result=[]
        urllist=['''http://f10.eastmoney.com/f10_v2/CompanySurvey.aspx?code=SZ''','''http://f10.eastmoney.com/f10_v2/CompanySurvey.aspx?code=SH''']
        for alistdata in listdata:
            tmpdata=[]
            for adata in alistdata:
                if adata[0][0]=='6':
                    href=urllist[1]+adata[0]
                else:
                    href=urllist[0]+adata[0]
                self.driver.get(href)
                errornum=0
                while 1:
                    errornum+=1
                    try:
                        industryNmae=self.driver.find_element_by_xpath("//*[@id=\"Table0\"]/tbody/tr[8]/td[2]").text
                        break
                    except:
                        if errornum%5==0:
                            self.driver.refresh()
                        time.sleep(0.5)
                        continue

                industryNmae=industryNmae.split("-")[-1]
                print(industryNmae)
                listedTime=self.driver.find_element_by_xpath("//*[@id=\"templateDiv\"]/div[2]/div[2]/table/tbody/tr[1]/td[2]").text

                adata[3]=industryNmae
                adata.append(listedTime)
                print(adata)
                tmpdata.append(adata)
            result.append(tmpdata)
        return result

    #获得一个榜单数据
    def getAbankuai(self):
        #点击涨幅榜单
        js='''document.querySelector("#dataview > div.dataview-center > div.dataview-body > table > thead > tr:nth-child(1) > th:nth-child(6) > div").click()'''
        self.driver.execute_script(js)
        time.sleep(6)
        up=self.getLeaderboarddata()
        self.driver.execute_script(js)
        time.sleep(6)
        down=self.getLeaderboarddata()
        if "-" in up[0]:#如何涨和跌位置反了
            tmp=down
            down=up
            up=tmp
        print(up)
        print(down)
        return [up,down]

    #获取点击榜单后的主要数据
    def getLeaderboarddata(self):
        content=self.driver.find_element_by_xpath("//*[@id=\"dataview\"]/div[2]/div[2]/table/tbody")
        trs=content.find_elements_by_tag_name("tr")
        targetdata=[]#目标数据
        for atr in trs:
            tmplist=atr.text
            tmplist=tmplist.split(" ")
            if tmplist[2][0]=="N":
                continue
            infoUrl=atr.find_elements_by_tag_name("a")[1].get_attribute("href")
            targetdata.append([tmplist[1],tmplist[2],tmplist[9],infoUrl])
            if len(targetdata)==10:
                break
        return targetdata

    #数据转化成可以直接写的数据
    def data2writdata(self,ndata):
        resdata=[]
        resdata.append( [" ","日期:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()), " ", " ", " "])
        resdata.append( ["涨幅榜"," ", " ", " ", " ", " "])
        adddFlag=True
        for adata in ndata:
            resdata.append([" 序号","代码", "名称", "涨跌幅", "所属行业", "上市时间"])
            for i,aadata in enumerate(adata):
                aadata.insert(0,str(i+1))
                resdata.append(aadata)
            if adddFlag:
                resdata.append( [" "," ", " ", " ", " ", " "],)
                resdata.append( ["跌幅榜"," ", " ", " ", " ", " "],)
                adddFlag=False
        return resdata
    #将数据写入xls
    def write_excel_xls(self,path,nameList,valueList):
        workbook = xlwt.Workbook()  # 新建一个工作簿
        for i,aname in enumerate(nameList):
            sheet_name=aname
            value=valueList[i]
            index = len(value)  # 获取需要写入数据的行数
            sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
            sheet.col(4).width = 8888
            for i in range(0, index):
                sheet.row(i).set_style(tall_style)
                for j in range(0, len(value[i])):
                    sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
        workbook.save(os.path.dirname(sys.executable)+"\\"+path)
        print("xls格式表格写入数据成功！")
def run():
    aobj=splider()
    aobj.major()
    aobj.driver.close()
    cmd = "taskkill /f  /im  chromedriver.exe  -T"
    res = subprocess.call(cmd, shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

if __name__ == '__main__':

    if len(sys.argv)==3 and sys.argv[1]=='-t' and sys.argv[2]!="":
        runtime=sys.argv[2]
        runtime=runtime.replace('：',':')
        while 1:
            nowtime=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())[11:16]
            print(runtime,nowtime)
            if runtime==nowtime:
                run()
                print("完成执行，等待下次执行")
                time.sleep(60*60*23)
            else:
                print("时间未到等待一会")
                time.sleep(10)
    else:
        run()
        sys.exit(0)




