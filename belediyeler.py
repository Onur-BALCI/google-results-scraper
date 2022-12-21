import re
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
from openpyxl import load_workbook


def find_email(i):
    while i >= 0:
        raw, email = "", ""
        driver.get(f'https://www.google.com/search?q={ara} url:{filt}')
        time.sleep(2)
        try:
            driver.find_element(By.ID, "L2AGLb").click()
        except selenium.common.exceptions.NoSuchElementException:
            pass
        try:
            driver.find_elements(By.TAG_NAME, 'h3')[i].click()
            time.sleep(5)
            raw = driver.find_element(By.XPATH, '/html/body').text
            email = (re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', raw))
        except:
            pass
        if i == 4:
            sheet[f"F{n+2}"] = str(email)
            sheet[f"G{n+2}"] = str(raw)
        if i == 3:
            sheet[f"H{n+2}"] = str(email)
            sheet[f"I{n+2}"] = str(raw)
        if i == 2:
            sheet[f"J{n+2}"] = str(email)
            sheet[f"K{n+2}"] = str(raw)
        if i == 1:
            sheet[f"L{n+2}"] = str(email)
            sheet[f"M{n+2}"] = str(raw)
        if i == 0:
            sheet[f"N{n+2}"] = str(email)
            sheet[f"O{n+2}"] = str(raw)
        i -= 1


if __name__ == '__main__':
    # <editor-fold desc="params">
    arama = pd.read_excel('input.xls', usecols='D', nrows=10)
    filtre = pd.read_excel('input.xls', usecols='E', nrows=10)
    arama = pd.DataFrame(arama).values.tolist()
    filtre = pd.DataFrame(filtre).values.tolist()
    driver = webdriver.Chrome()
    raws, ress = "", ""
    workbook = load_workbook(filename="input.xlsx")
    sheet = workbook.active
    # </editor-fold>
    for n, (ara, filt) in enumerate(zip(arama, filtre)):
        find_email(4)
    workbook.save(filename="output.xlsx")
