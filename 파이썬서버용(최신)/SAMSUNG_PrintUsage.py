from collections import defaultdict
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
import csv
from datetime import datetime
import time
import PrintUsageDB as db
#from pandas import Series, DataFrame
import smtplib
from email import encoders
from email.utils import formataddr
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
 #sendmail.py
import os
import smtplib

#에몬스홈 = 206
#개발실 = 220
#전시장 = 214

Device = {'에몬스앳홈' : 206,'개발실':220}

      
findDept = ''

db.dbConnect()
for Di in Device:
    rdr = {}
    rdr = db.selectIndividual(Di)
#f = open('StandardAccount.csv','r')
#rdr = csv.DictReader(f)


    

    '''START웹 크롤링 사이트 연결'''
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
    driver.implicitly_wait(5) # 암묵적으로 웹 자원을 (최대) 3초 기다리기
    # Login
    DeviceDriver = 'http://192.168.0.{}/sws.login/gnb/loginView.sws?basedURL=undefined&popupid=id_Login'.format(Device.get(Di))
    driver.get(DeviceDriver) 
    driver.find_element_by_name('IDUserId').send_keys('admin') # 값 입력
    driver.find_element_by_name('IDUserPw').send_keys('soa4007@')
    driver.find_element_by_name('IDUserPw').send_keys(Keys.ENTER)

    '''웹 크롤링 사이트 연결 END'''

    '''START 흑백, 컬러 총 합 매수 출력'''
    timestr = time.strftime("%Y%m%d-%H%M%S")
    filedName = ['흑백총합','컬러총합']
    dictWrite = open(timestr+'TOTAL'+'.csv','a',newline="")
    dictWrite = csv.DictWriter(dictWrite,fieldnames=filedName)

    UsePrintTotalSite = 'http://192.168.0.{}/sws.application/information/countersView.sws?ruiFw_id=Counters&ruiFw_pid=Information&ruiFw_title=%EC%82%AC%EC%9A%A9%20%EC%B9%B4%EC%9A%B4%ED%84%B0'.format(Device.get(Di))
    driver.get(UsePrintTotalSite)
    html = driver.page_source # 페이지의 elements모두 가져오기
    soup = BeautifulSoup(html, 'html.parser') # BeautifulSoup사용하기
    temp = soup.find('table',{'id':'swstable_counterTotalList_contentTB'})
    temp_td = temp.find_all('td')
    temp_tdList = []

    temp_tdLen= len(temp_td)

    for i in range(0,temp_tdLen):
        temp_td = temp.find_all('td')[i].text
        temp_tdList.append(temp_td)
        

    print(temp_tdList[9]) #흑백 총 매수
    print(temp_tdList[19]) # 컬러 총 매수 

    monoTotalUse = temp_tdList[9]
    colorTotalUse = temp_tdList[19]

    db.dbInsertTotal(datetime.today().strftime("%Y%m%d"),monoTotalUse,colorTotalUse,Di,datetime.today())
    dictWrite.writerow({'흑백총합': monoTotalUse,'컬러총합':colorTotalUse})
    '''흑백, 컬러 총 합 매수 출력 END '''





    '''START 개인별 프린터 출력 / 복사 매수 출력'''
    timestr = time.strftime("%Y%m%d-%H%M%S")
    value={}
    filedName = ['부서','성명','사번','컬러출력','흑백출력','컬러복사','흑백복사']
    dictWrite = open(timestr+'.csv','a',newline="")
    dictWrite = csv.DictWriter(dictWrite,fieldnames=filedName)
    dictWrite.writeheader()


    for i in rdr:
        
        #UsePrintSite = 'http://192.168.0.{}/sws.application/userManagement/editAccountID.sws?userid={}&popupid=editAccountID'.format(Device.get(Di),i[2])
        UsePrintSite = 'http://192.168.0.{}/sws.application/userManagement/editAccountID.sws?userid={}&popupid=editAccountID'.format(Device.get(Di),i['UserID'])

        driver.get(UsePrintSite)
        html = driver.page_source # 페이지의 elements모두 가져오기

        soup = BeautifulSoup(html, 'html.parser') # BeautifulSoup사용하기

        colorPrintUse = soup.find('td',{'id':'colorPrintUse'}).text
        monoPrintUse = soup.find('td',{'id':'monoPrintUse'}).text
        
        colorCopyUse = soup.find('td',{'id':'colorCopyUse'}).text
        monoCopyUse = soup.find('td',{'id':'monoCopyUse'}).text
        
        #monoCopyUse
        #colorCopyUse

        
        #db.dbExcute(db.FindQuery(Di,i[2]))
        db.dbExcute(db.FindQuery(Di,i['UserID']))
        DeptandUserName = db.dbFetchOne()
        while DeptandUserName:
            #print("부서" + str(DeptandUserName[0]) + ", 성명 " + str(DeptandUserName[1]))
            DeptName = str(DeptandUserName['DeptName'])
            UserName = str(DeptandUserName['EmpName'])
            #print(DeptName + UserName)
            DeptandUserName = db.dbFetchOne()
            
            db.dbInsertIndividual(datetime.today().strftime("%Y%m%d"),i['UserID'],DeptName,UserName,monoPrintUse,colorPrintUse,monoCopyUse,colorCopyUse,Di,datetime.today())
            dictWrite.writerow({'부서':DeptName,'성명': UserName,'사번' : i['UserID'], '컬러출력' : colorPrintUse,
                                '흑백출력' : monoPrintUse, '컬러복사' : colorCopyUse, '흑백복사' : monoCopyUse })

            #raw_data = {'부서':DeptName,'성명': UserName,'사번' : i['AcountID'], '컬러출력' : colorPrintUse, '흑백출력' : monoPrintUse}
            #dataFrame = DataFrame(raw_data,index=[i])
            #print(dataFrame)
      
    '''개인별 프린터 출력 / 복사 매수 출력 END'''
    driver.quit()

driver.quit()
db.dbClose()

