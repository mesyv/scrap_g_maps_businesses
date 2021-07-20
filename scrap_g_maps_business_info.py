#Importing libraries 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
from time import sleep
from datetime import datetime, date
import openpyxl
import time 
import string
import openpyxl
import os

#Chromium setting
options = Options()
options.add_argument(r"--user-data-dir=C:\Users\kamil.smolag\AppData\Local\Google\Chrome\User Data")
options.add_argument(r"--profile-directory=Profile 5")
driver = webdriver.Chrome(executable_path=r"C:\Projects\chromedriver.exe", chrome_options=options)

#zmienne
waiting_time = 2
time_today = date.today()
i = 1
ID = 2
page = 0
workbook = Workbook()
worksheet = workbook.active
worksheet["A1"] = "ID"
worksheet["B1"] = "miejsce"
worksheet["C1"] = "miasto"
worksheet["D1"] = "nazwa"
worksheet["E1"] = "opinie"
worksheet["F1"] = "srednia"
worksheet["G1"] = "adres"
worksheet["H1"] = "strona"
worksheet["I1"] = "telefon"
worksheet["J1"] = "URL"
worksheet["K1"] = " "
miejsce = "Dentist NYC"
miasto = " "

driver.get("https://www.google.pl/maps")
sleep(waiting_time)
place = driver.find_element_by_class_name("tactile-searchbox-input")
place.send_keys(f"{miejsce} {miasto}")
place.send_keys(Keys.RETURN)
sleep(waiting_time)

for page in range(5):
    while i < 50:
        for page_one in range(page):
            next_page = driver.find_element_by_xpath("//button[@id='ppdPk-Ej1Yeb-LgbsSe-tJiF1e']")
            next_page.click()
            sleep(2)
        #scroll na wynikach wyszukiwania
        try:
            scroll_wynik = driver.find_element_by_xpath("/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]/div/div/div[4]/div[1]/div[3]/div/a")
            for x in range(40):
                scroll_wynik.send_keys(Keys.PAGE_DOWN)
        except:
            pass

        #klik na lokacje
        try:
            loc_1 = driver.find_element_by_xpath(f"/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]/div/div/div[4]/div[1]/div[{i}]/div/a")
            driver.execute_script("arguments[0].scrollIntoView();", loc_1)
            loc_1.click()
            location1 = True
        except:
            print("No loc click")
            location1 = False
            pass
        
        sleep(waiting_time)

        #nazwa wizytówki
        try:
            nazwa = driver.find_element_by_xpath("/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]").text
        except:
            nazwa = "N/a"
            pass
        #ilosc opinii
        try:
            opinie = driver.find_element_by_xpath("/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[2]/div/div[1]/span[1]/span/span/span[2]/span[1]/button").text
        except:
            opinie = "N/a"
            pass
        #srednia ocen
        try:
            srednia = driver.find_element_by_xpath("/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[2]/div/div[1]/div[2]/span/span/span").text
        except:
            srednia = "N/a"
            pass
        #scroll na wizytowce
        try:
            scroll = driver.find_element_by_xpath("/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]")
            scroll.send_keys(Keys.PAGE_DOWN)
        except:
            pass
        #adres
        try:
            adres = driver.find_element_by_xpath("//button[@class='CsEnBe' and @data-item-id='address']").text
        except NoSuchElementException:
            adres = "N/a"
            pass
        #strona
        try:
            strona = driver.find_element_by_xpath("//button[@class='CsEnBe' and @data-item-id='authority']").text
        except NoSuchElementException:
            strona = "N/a"
            pass
        #telefon
        try:
            telefon = driver.find_element_by_xpath("//button[@class='CsEnBe' and @data-tooltip='Kopiuj numer telefonu']").text
        except NoSuchElementException:
            telefon = "N/a"
            pass
        
        #zapis do pliku
        if nazwa == "N/a":
            print("Pass")
        else:
            worksheet[f"A{ID}"] = ID
            worksheet[f"B{ID}"] = miejsce
            worksheet[f"C{ID}"] = miasto
            worksheet[f"D{ID}"] = nazwa
            worksheet[f"E{ID}"] = opinie
            worksheet[f"F{ID}"] = srednia
            worksheet[f"G{ID}"] = adres
            worksheet[f"H{ID}"] = strona
            worksheet[f"I{ID}"] = telefon
            worksheet[f"J{ID}"] = driver.current_url
            worksheet[f"K{ID}"] = " "
            workbook.save(f"scrap_{miejsce}_{time_today}.xlsx")
            ID += 1

            print(f"Pętla nr: {ID}")
            print(driver.current_url)
            print(nazwa)
            print(opinie)
            print(srednia)
            print(adres)
            print(strona)
            print(telefon)
            print(" ")

        print(f"i: {i}\n")
        print(f"page: {page}\n")
        i += 2
        if location1 == True:
            driver.back()
            sleep(waiting_time)
    page += 1
    i = 1

print("End")
driver.quit()
