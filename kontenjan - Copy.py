import mechanize
from bs4 import BeautifulSoup
import openpyxl
import time

import ctypes

wb = openpyxl.Workbook()
ws = wb.active

print("Course Capacity Checker Created by Ali Yilbasi & Ahmet Bakkal 2022\n")
ders = input("Course Code (i.e. END458E):").upper()
crn = input("CRN Number (i.e. 21268):")

def itukont(ders,crn):

    browser = mechanize.Browser()
    browser.set_handle_robots(False)
    browser.open("https://www.sis.itu.edu.tr/TR/ogrenci/ders-programi/ders-programi.php?seviye=LS")
    browser.select_form(nr=0)

    browser.form["derskodu"]=[ders[:3]]

    browser.submit()
    soup_table=BeautifulSoup(browser.response().read(),'html.parser')
    table = soup_table.find('table')

    a = 0
    c = 0
    for i in table.find_all('tr'):
        b = 0
        c += 1
        if c == 1 or c == 2:
            pass
        else:
            a += 1
        for j in i.find_all('td'):
            b += 1
            if c == 1 or c == 2:
                pass
            else:
                ws.cell(column=b, row=a).value = j.get_text(strip=True)

    for row in ws.rows:
        if row[0].value == crn:
            for cell in row:
                if int(row[9].value) > int(row[10].value):
                    check = "Course is available"
                    ctypes.windll.user32.MessageBoxW(0,"Course is available, run forest run :)", "Attention from Course Capacity Checker",0x40000)
                    print(check)
                    break
                else:
                    check = "Course is full"
                    print(check)

                    break


while True:
   itukont(ders,crn)
   time.sleep(5)
