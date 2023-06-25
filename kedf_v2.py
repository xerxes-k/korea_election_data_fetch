import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select


########## version_2 SE + BS4

################# 사이트 로딩
driver = webdriver.Chrome()
driver.implicitly_wait(3)
driver.get("http://info.nec.go.kr/main/showDocument.xhtml?electionId=0000000000&topMenuId=CP&secondMenuId=CPRI03")
#혹시 모르니 창 최대화
driver.maximize_window()

################# 엑셀을 만든다
wb = Workbook(write_only=True)
ws = wb.create_sheet('election')
ws.append(['region', 'party', 'name', 'sex', 'dateofbirth', 'occupation', 'education', 'etc'])


################# 옵션 설정을 위한 셀렉트 준비

congress = driver.find_element(By.CSS_SELECTOR, '#electionType2')
year = driver.find_element(By.CSS_SELECTOR, '#electionName')
election = driver.find_element(By.CSS_SELECTOR, '#electionCode')
code = driver.find_element(By.CSS_SELECTOR, '#cityCode')
btn = driver.find_element(By.CSS_SELECTOR, '#searchBtn')
congress.click()

# 연도 마다 돌아가면서
for i in range(1, 4):
    Select(year).select_by_index(i)
    WebDriverWait(driver, 3).until(EC.visibility_of(election))
    Select(election).select_by_value('2')
    WebDriverWait(driver, 3).until(EC.visibility_of(code))
    Select(code).select_by_index(1)
    btn.click()
    year = driver.find_element(By.CSS_SELECTOR, '#electionName')
    election = driver.find_element(By.CSS_SELECTOR, '#electionCode')
    code = driver.find_element(By.CSS_SELECTOR, '#cityCode')
    btn = driver.find_element(By.CSS_SELECTOR, '#searchBtn')
    
    # html을 bs4로 파싱하기
    elec_page = driver.page_source
    soup = BeautifulSoup(elec_page, 'html.parser')
    trs = soup.select('tbody tr')
    for tr in trs:
        tds = tr.select('td')
        lt = []
        for td in tds:
            lt.append(td.get_text())
        # print(lt)
        # reg = lt[0]
        # par = lt[2]
        # name = lt[3]
        # sex = lt[4]
        # dob = lt[5]
        # occ = lt[6]
        # edu = lt[7]
        # etc = lt[8]
        ws.append(lt)

wb.save('election.xlsx')        

driver.quit()