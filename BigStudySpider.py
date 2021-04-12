#author：袁守航
from bs4 import BeautifulSoup
import re
import xlwt
import xlrd
import tkinter
namelist = xlrd.open_workbook('云上青春录入团支部.xls')
from selenium import webdriver
findLink = re.compile(r'<p>(.*?)</p>')     #创建正则表达式，表示规则
i = 0

url = 'http://admin.ddy.tjyun.com'
username = ''      #账号
password = ''           #密码
workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
worksheet = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
sheet = namelist.sheet_by_index(0)
driver = webdriver.Chrome()
driver.get(url)
driver.find_element_by_class_name('jinyun_control_loginblock_lname').click()
driver.find_element_by_id('userName').send_keys(username)
driver.find_element_by_class_name('jinyun_control_loginblock_lpassword').click()
driver.find_element_by_id('password').send_keys(password)
VerificationCode = input('请输入验证码：')
driver.find_element_by_class_name('input-text').click()
driver.find_element_by_id('imageCode').send_keys(VerificationCode)
driver.find_element_by_class_name('jinyun_control_loginblock_lenter').click()
driver.find_element_by_link_text('大学习统计').click()

for i in range(0,53):
    j = 1
    classNameSimple = sheet.cell_value(i,1)
    classNameComplex = sheet.cell_value(i,2)
    str(classNameSimple)
    str(classNameComplex)
    driver.find_element_by_link_text('大学习统计').click()
    driver.find_element_by_id('tree1').click()
    try:
        driver.find_element_by_link_text(classNameComplex).click()
    except:
        worksheet.write(i, 0, classNameSimple)
        continue
    soup = BeautifulSoup(driver.page_source,'html.parser')
    worksheet.write(i,0,classNameSimple)
    for item in soup.find_all('div', class_="col-xs-2"):
        data = []
        item = str(item)
        name = re.findall(findLink, item)[0]
        data.append(name)
        worksheet.write(i, j, data)
        j += 1
        workbook.save('青年大学习统计.xls')
        continue
