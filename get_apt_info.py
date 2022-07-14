from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import openpyxl
import os
"""
Date : 2020/05/05
Author : 이충영
Description : Naver에서 단지 검색 후, 필요한 정보(단지명, 세대수, 준공일 등) Excel 파일로 추출
주의사항 : Chromedriver는 설치된 Chrome의 버전과 동일해야함. 현 ChromeDriver 83.0.4103.39
"""
apt_name_file_addr = 'address.xlsx'
apt_name_excel = openpyxl.load_workbook(apt_name_file_addr)
apt_name_sheet = apt_name_excel.worksheets[0]
apt_name = []

col = apt_name_sheet['A']

for cell in col:
    apt_name.append(cell.value)

driver = Chrome(executable_path='chromedriver.exe')
time.sleep(0.1)
driver.get("https://www.naver.com")
time.sleep(0.1)

i = 0
for name in apt_name:
    i = i + 1
# 아파트 단지명 검색
    try:
        driver.find_element_by_id("query").send_keys(name+Keys.ENTER)
    except:
        driver.find_element_by_id("nx_query").send_keys(name+Keys.ENTER)
    time.sleep(0.1)

# 단지 세대수, 준공일, 주소 찾기
    try:
        building_info = driver.find_elements_by_xpath(
            "//*[@class = 'info_area']")[0].text
        list_building_info = building_info.split(' ')
        building_type = list_building_info[0]
        if building_type == "아파트분양권":
            building_type = "아파트"
        building_num = str(list_building_info[1]).replace("세대", "")
        building_date = list_building_info[3]
        print(str(building_type)+"  "+str(building_num)+"  "+str(building_date))
        building_addr = driver.find_elements_by_xpath(
            "//*[@class = 'addr']")[0].text
        print(str(building_addr))

    except:
        building_type = "wrong building name"
        building_num = "wrong building name"
        building_date = "wrong building name"
        building_addr = "wrong building name"

# 검색어창 Clear
    try:
        driver.find_element_by_id("query").clear()
    except:
        driver.find_element_by_id("nx_query").clear()

# Excel 파일에 입력
    apt_name_sheet['B'+str(i)] = building_num
    apt_name_sheet['C'+str(i)] = building_date
    apt_name_sheet['D'+str(i)] = building_addr
    apt_name_sheet['E'+str(i)] = building_type

# 바탕화면에 excel 파일 생성
apt_name_sheet.insert_rows(1)
apt_name_sheet['A1'] = '단지명'
apt_name_sheet['B1'] = '세대수'
apt_name_sheet['C1'] = '준공년월'
apt_name_sheet['D1'] = '주소'
apt_name_sheet['E1'] = '타입'
excel_addr = os.path.join(os.path.join(
    os.environ['USERPROFILE'])) + '\Desktop\아파트정보.xlsx'
apt_name_excel.save(excel_addr)
print(str(excel_addr))
print('----------------------end------------------------')
