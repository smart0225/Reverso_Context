from selenium import webdriver
import time
import datetime
import re
import pandas as pd
import numpy as np
#import win32com.client as win32
import os
import xlsx_py
#import openpyxl

#設定目前位置+存放EXCEL資料夾
start_Path = os.getcwd()
xl_Path = start_Path + "/xl"

#如果不存在
if not os.path.isdir(xl_Path):
    os.mkdir(xl_Path)

now_date = datetime.datetime.now().strftime('%Y%m%d')
file_path = xl_Path + '\\' + now_date + "_salebalancerank.xlsx"

#刪除已存在的檔案
try:
    os.remove(file_path)
except OSError as e:
    print(e)
else:
    print("File is deleted successfully.......")

#建立xlsx檔案
xlsx_py.buildxlsx(file_path)

chrome_options = webdriver.ChromeOptions()
#chrome_options.add_argument('--headless') # 啟動無頭模式
#chrome_options.add_argument('--disable-gpu') # windowsd必須加入此行
#chrome=webdriver.Chrome(chrome_options=chrome_options, executable_path='./chromedriver')
#chrome = webdriver.Chrome('chromedriver',chrome_options=chrome_options)
#chrome = webdriver.Chrome('chromedriver')

# 添加UA
#chrome_options.add_argument('user-agent="MQQBrowser/26 Mozilla/5.0 (Linux; U; Android 2.3.7; zh-cn; MB200 Build/GRJ22; CyanogenMod-7) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1"')
# 指定浏览器分辨率
#chrome_options.add_argument('window-size=1920x3000')

# 谷歌文档提到需要加上这个属性来规避bug
#chrome_options.add_argument('--disable-gpu')
#隐藏滚动条, 应对一些特殊页面
#chrome_options.add_argument('--hide-scrollbars')
# 不加载图片, 提升速度
#chrome_options.add_argument('blink-settings=imagesEnabled=false')
# 浏览器不提供可视化页面. linux下如果系统不支持可视化不加这条会启动失败
#chrome_options.add_argument('--headless')
# 以最高权限运行
chrome_options.add_argument('--no-sandbox')
# 手动指定使用的浏览器位置
chrome_options.binary_location = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
#添加crx插件
#option.add_extension('d:\crx\AdBlock_v2.17.crx')
# 禁用JavaScript
#chrome_options.add_argument("--disable-javascript")
# 设置开发者模式启动，该模式下webdriver属性为正常值
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
#反反爬蟲語法
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
# 禁用瀏覽器弹窗
#prefs = {'profile.default_content_setting_values':  {'notifications': 2 }}
#chrome_options.add_experimental_option('prefs',prefs)

chrome=webdriver.Chrome(chrome_options=chrome_options)

chrome.get("https://www.wantgoo.com/stock/margin-trading/short-lending-sale-balance-rank")
time.sleep(2)
#關閉廣告
#try:
#    suspondwindows = chrome.find_element_by_xpath('//*[@id="popupTest"]/div/div/div/div/div/div[1]/div/div[3]/a')
#    suspondwindows.click()
#    print('close ok!!')
#except Exception as e:
#    print('close ng!!')

#time.sleep(2)
#讀取網頁CLASS相關字
datas = chrome.find_elements_by_class_name('rt')
time.sleep(1.5)
#(上市櫃)借券賣出餘額增加
arrPair = re.findall(r'\S+',datas[0].text)
grid_txt = "_".join(arrPair)
time.sleep(1.5)
n=len(grid_txt.split('_'))/5
result = list(np.array_split(grid_txt.split('_'),n))
#list轉換成DataFrame格式
df = pd.DataFrame(result,columns=['排名','股票','當日餘額','增減張數','增減金額(萬)'])
df.sort_values(by=['增減金額(萬)'], inplace=True,ascending=[True])

time.sleep(1.5)

#借券賣出餘額減少
arrPairA = re.findall(r'\S+',datas[1].text)
gridA_txt = "_".join(arrPairA)
time.sleep(1.5)
n1=len(gridA_txt.split('_'))/5
resultA = list(np.array_split(gridA_txt.split('_'),n1))
#list轉換成DataFrame格式
dfA = pd.DataFrame(resultA,columns=['排名','股票','當日餘額','增減張數','增減金額(萬)'])
dfA.sort_values(by=['增減金額(萬)'], inplace=True,ascending=[True])

#20210722新增加爬'(上櫃)借券賣出餘額增減排行'
chrome.get('https://www.wantgoo.com/stock/margin-trading/short-lending-sale-balance-rank?market=OTC')
datasB = chrome.find_elements_by_class_name('rt')
time.sleep(1.5)
#借券賣出餘額增加
arrPairB = re.findall(r'\S+',datasB[0].text)
grid_txtB = "_".join(arrPairB)
time.sleep(1.5)
n2=len(grid_txtB.split('_'))/5
resultB = list(np.array_split(grid_txtB.split('_'),n2))
#list轉換成DataFrame格式
dfB = pd.DataFrame(resultB,columns=['排名','股票','當日餘額','增減張數','增減金額(萬)'])
dfB.sort_values(by=['增減金額(萬)'], inplace=True,ascending=[True])

time.sleep(1.5)
#借券賣出餘額減少
arrPairC = re.findall(r'\S+',datasB[1].text)
grid_txtC = "_".join(arrPairC)
time.sleep(1.5)
n3=len(grid_txtC.split('_'))/5
resultC = list(np.array_split(grid_txtC.split('_'),n3))
#list轉換成DataFrame格式
dfC = pd.DataFrame(resultC,columns=['排名','股票','當日餘額','增減張數','增減金額(萬)'])
dfC.sort_values(by=['增減金額(萬)'], inplace=True,ascending=[True])

#20210726(上市)借券賣出餘額增減排行
chrome.get('https://www.wantgoo.com/stock/margin-trading/short-lending-sale-balance-rank?market=Listed')
datasD = chrome.find_elements_by_class_name('rt')
time.sleep(1.5)
#借券賣出餘額增加
arrPairD = re.findall(r'\S+',datasD[0].text)
grid_txtD = "_".join(arrPairD)
time.sleep(1.5)
n4=len(grid_txtD.split('_'))/5
resultD = list(np.array_split(grid_txtD.split('_'),n4))
#list轉換成DataFrame格式
dfD = pd.DataFrame(resultD,columns=['排名','股票','當日餘額','增減張數','增減金額(萬)'])
dfD.sort_values(by=['增減金額(萬)'], inplace=True,ascending=[True])

#借券賣出餘額減少
arrPairE = re.findall(r'\S+',datasD[1].text)
grid_txtE = "_".join(arrPairE)
time.sleep(1.5)
n5=len(grid_txtE.split('_'))/5
resultE = list(np.array_split(grid_txtE.split('_'),n5))
#list轉換成DataFrame格式
dfE = pd.DataFrame(resultE,columns=['排名','股票','當日餘額','增減張數','增減金額(萬)'])
dfE.sort_values(by=['增減金額(萬)'], inplace=True,ascending=[True])

#結束爬蟲，關閉chrome物件
chrome.quit()
time.sleep(1.5)

#因EXCEL版本不同，只好運用numpy轉成一個EXCEL檔案在OPEN處理(此為Numpy方式寫EXCEL語法)
#df.to_excel(file_path,'N',columns=['排名','股票','當日餘額','增減張數','增減金額(萬)'],index=False)

#使用Pandas開始將資料寫入EXCEL
with pd.ExcelWriter(file_path) as writer:
    df.to_excel(writer,sheet_name="(上市櫃)借劵賣出餘額增加",index=False,float_format="%0.1f")
    dfA.to_excel(writer,sheet_name="(上市櫃)借劵賣出餘額減少",index=False,float_format="%0.1f")
    dfB.to_excel(writer,sheet_name="(上櫃)借券賣出餘額增加",index=False,float_format="%0.1f")
    dfC.to_excel(writer,sheet_name="(上櫃)借券賣出餘額減少",index=False,float_format="%0.1f")
    dfD.to_excel(writer,sheet_name="(上市)借券賣出餘額增加",index=False,float_format="%0.1f")
    dfE.to_excel(writer, sheet_name="(上市)借券賣出餘額減少", index=False,float_format="%0.1f")

time.sleep(1.5)

#修改EXCEL儲存格
xlsx_py.Fix_Cells(file_path)

#xls_path = xl_Path[0:2] + xl_Path[3:] + '\\' + now_date + "_salebalancerank.xlsx"
#str = ['排名','股票','當日餘額','增減張數','增減金額(萬)']
#此為win32方式控制EXCEL，此方式有個缺點，會在USER資料夾\AppData\Local\Temp\gen_py\3.8(版本)下新增一些物件導致某次使用異常
#ExcelApp = win32.gencache.EnsureDispatch('Excel.Application')
#ExcelApp.Visible = False
#ExcelApp.DisplayAlerts = False
# newBook = ExcelApp.Workbooks.Open(file_path)
# sheet = newBook.Worksheets.Add()
# sheet.Name = '(上市櫃)借劵賣出餘額增加'
# sheet.Range(sheet.Cells(1,1),sheet.Cells(1,5)).Value = str
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(df.index)-1,1+len(df.columns)-1)).Value = df.values
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(df.index)-1,1+len(df.columns)-1)).Sort(Key1=sheet.Range('E1'),Order1=2, Orientation=1)
# sheet = newBook.Worksheets.Add()
# sheet.Name = '(上市櫃)借劵賣出餘額減少'
# sheet.Range(sheet.Cells(1,1),sheet.Cells(1,5)).Value = str
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfA.index)-1,1+len(dfA.columns)-1)).Value = dfA.values
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfA.index)-1,1+len(dfA.columns)-1)).Sort(Key1=sheet.Range('E1'),Order1=2, Orientation=1)
# #20210722
# sheet = newBook.Worksheets.Add()
# sheet.Name = '(上櫃)借券賣出餘額增加'
# sheet.Range(sheet.Cells(1,1),sheet.Cells(1,5)).Value = str
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfB.index)-1,1+len(dfB.columns)-1)).Value = dfB.values
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfB.index)-1,1+len(dfB.columns)-1)).Sort(Key1=sheet.Range('E1'),Order1=2, Orientation=1)
# sheet = newBook.Worksheets.Add()
# sheet.Name = '(上櫃)借券賣出餘額減少'
# sheet.Range(sheet.Cells(1,1),sheet.Cells(1,5)).Value = str
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfC.index)-1,1+len(dfC.columns)-1)).Value = dfC.values
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfC.index)-1,1+len(dfC.columns)-1)).Sort(Key1=sheet.Range('E1'),Order1=2, Orientation=1)
# #20210726
# sheet = newBook.Worksheets.Add()
# sheet.Name = '(上市)借券賣出餘額增加'
# sheet.Range(sheet.Cells(1,1),sheet.Cells(1,5)).Value = str
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfD.index)-1,1+len(dfD.columns)-1)).Value = dfD.values
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfD.index)-1,1+len(dfD.columns)-1)).Sort(Key1=sheet.Range('E1'),Order1=2, Orientation=1)
# sheet = newBook.Worksheets.Add()
# sheet.Name = '(上市)借券賣出餘額減少'
# sheet.Range(sheet.Cells(1,1),sheet.Cells(1,5)).Value = str
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfE.index)-1,1+len(dfE.columns)-1)).Value = dfE.values
# sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfE.index)-1,1+len(dfE.columns)-1)).Sort(Key1=sheet.Range('E1'),Order1=2, Orientation=1)

#sheet = newBook.Worksheets().Add()
#sheet.Name = '借劵賣出餘額減少'
#sheet.Range(sheet.Cells(1,1),sheet.Cells(1,5)).Value = str
#sheet.Range(sheet.Cells(2,1),sheet.Cells(2+len(dfA.index)-1,1+len(dfA.columns)-1)).Value = dfA.values
#因EXCEL版本不同，開啟時有些有Sheet 3個，新版只有一個
#newBook.Worksheets('N').Delete()

##newBook.SaveAS("xx.xlsx")
#newBook.Close(True)

time.sleep(1.5)

#套用C#建置Kill excel處理程序功能
#kill_xl_Path = start_Pathrw_xlsx + '\\Kill.exe'
#os.system(kill_xl_Path)
#print(kill_xl_Path)

#kill excel及chrome處理程序(取代之前外掛模式，直接於"處理程序"刪除)
xlsx_py.Kill_XLSX()

time.sleep(2)
print("所有程序已執行完畢~~~~~~")