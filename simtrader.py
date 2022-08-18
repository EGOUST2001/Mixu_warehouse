import datetime
import os
import time
import re
import numpy as np
import pandas as pd
from urllib.request import urlretrieve
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from PIL import Image, ImageEnhance, ImageFilter
from PIL.ImageFilter import Color3DLUT
from pathlib import Path
from styleframe import StyleFrame, Styler, utils
import xlrd
import xlwt
import openExcel
import openpyxl as ox


# def get_c():
#     s = Service('yandexdriver.exe')
#     driver = webdriver.Chrome(service=s)
#     url = "https://dss.simtrader.ru/modules/auth.php"
#     driver.get(url)
#
#     driver.save_screenshot("crop_captcha.png")
#     img = Image.open('crop_captcha.png')
#     img_width, img_height = img.size
#     crop_img = img.crop((img_width/1.9, img_height/1.9, img_width/1.7, img_height/1.7))
#     crop_img.save('crop_captcha.png', quality=95)
#     return driver
def get_c():
    s = Service('chromedriver.exe')
    driver = webdriver.Chrome(service=s)
    url = "https://dss.simtrader.ru/modules/auth.php"
    driver.get(url)
    driver.fullscreen_window()
    driver.find_element(by=By.TAG_NAME, value="img").screenshot('crop_captcha.png')
    driver.minimize_window()
    return driver


# def get_c():
#     s = Service('yandexdriver.exe')
#     driver = webdriver.Chrome(service=s)
#     url = "https://dss.simtrader.ru/modules/auth.php"
#     driver.get(url)
#     driver.execute_script("document.body.style.zoom='100%'")
#     driver.fullscreen_window()
#     driver.find_element(by=By.TAG_NAME, value="img").screenshot("crop_captcha.png")
#     driver.minimize_window()
#     return driver


# вход в simtrade
def log_in_to_the_system(driver, login, password, cap):
    el_c = driver.find_elements(by=By.XPATH, value="//input[@name='captcha']")

    el_c[0].send_keys(cap)

    el_login = driver.find_elements(by=By.XPATH, value="//input[@name='login']")
    el_login[0].send_keys(login)

    el_password = driver.find_elements(by=By.XPATH, value="//input[@name='pass']")
    el_password[0].send_keys(password)

    btn_entr = driver.find_elements(by=By.XPATH, value="//input[@type='submit']")
    btn_entr[0].click()

    # return driver


# получение всех операторов
def get_operator(driver):
    btn_entrance = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Операторы')]")
    btn_entrance[0].click()

    operators = []
    for i in range(2, 7):
        c = driver.find_element(by=By.XPATH, value=f"//*[@class='i_list']/tbody/tr[{i}]/td[3]")
        operators.append(c.text)

    return operators


# Получении словаря (Тариф: Оператор)
def get_rates(driver, operators):
    btn_entrance = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Тарифные планы')]")
    btn_entrance[0].click()

    rates = {}
    for i in operators:
        select = Select(driver.find_elements(by=By.XPATH, value="//select[@name='operatorid']")[0])
        select.select_by_visible_text(i)

        r = driver.find_elements(by=By.XPATH, value="//table[@class= 'i_list']/tbody/tr")
        for j in range(2, len(r)):
            c = driver.find_element(by=By.XPATH, value=f"//*[@class='i_list']/tbody/tr[{j}]/td[3]")
            rates[c.text] = i

    return rates


# Поступление
def entrance(driver, path):  # rates,
    try:
        # sim_reconciliation(driver, path)
        df = get_exl(path)
        dop_info = False
        if df.shape[1] == 5:
            dop_info = True
        if df.shape[1] == 5 or df.shape[1] == 4:
            Operators = df[2].unique()
            for oper in Operators:
                new_df = df[df[2] == oper]
                tarifs = new_df[3].unique()
                for tarif in tarifs:
                    dop_df = new_df[new_df[3] == tarif]
                    dop_df.reset_index(inplace=True)
                    nots = ""
                    if dop_info:
                        dop_df[4].fillna("", inplace=True)
                        nots = dop_df[4][0]
                    dop_df = dop_df[[0, 1]]
                    dop_df[0].fillna(dop_df[1], inplace=True)
                    dop_df.to_excel("доп/д_реестр.xlsx", index=False, header=None)
                    inExcel = (r"доп/д_реестр.xlsx")
                    workbook = xlrd.open_workbook(inExcel)
                    sheetIn = workbook.sheet_by_index(0)
                    workbook = xlwt.Workbook()
                    sheetOut = workbook.add_sheet('DATA')
                    for r in range(sheetIn.nrows):
                        for c in range(sheetIn.ncols):
                            sheetOut.write(r, c, sheetIn.cell_value(r, c))

                    workbook.save(inExcel)
                    btn_entrance = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Поступление СИМ')]")
                    btn_entrance[0].click()
                    btn_new_entrance = driver.find_elements(by=By.XPATH, value="//*[contains(text(), 'Новый')]")
                    btn_new_entrance[0].click()

                    select1 = Select(driver.find_elements(by=By.XPATH, value="//select[@name='operatorid']")[0])
                    select1.select_by_visible_text(oper)  # rates.get(t).rstrip().lstrip()
                    inp_note = driver.find_element(by=By.XPATH, value="//input[@name='note']")
                    inp_note.send_keys(nots)
                    btn_add2 = driver.find_elements(by=By.XPATH, value="//input[@value='Записать']")
                    btn_add2[0].click()
                    time.sleep(1)
                    select2 = Select(driver.find_elements(by=By.XPATH, value="//select[@name='tplanid']")[0])
                    select2.select_by_visible_text(tarif)

                    field = driver.find_element(by=By.NAME, value="filename")
                    field.send_keys(os.getcwd() + "/доп/д_реестр.xlsx")

                    btn_add = driver.find_elements(by=By.XPATH, value="//input[@value='Добавить']")
                    btn_add[0].click()

                    time.sleep(1)
                    btn_add3 = driver.find_element(by=By.XPATH, value="//input[@value='OK']")
                    btn_add3.click()
        else:
            from main import ErrorDialog
            er = ErrorDialog("Неверный формат файла. Файл должен содержать 4 или 5 столбцов")
            er.exec()
    except Exception as e:
        from main import ErrorDialog
        if str(e) == "list index out of range":
            er = ErrorDialog("Неверный логин, пароль или капча")
            er.exec()
        if "Could not locate" in str(e):
            er = ErrorDialog("Такого тарифа или оператора не существует: " +
                             str(e).split()[-1])
            er.exec()
        else:
            er = ErrorDialog(str(e))
            er.exec()


# сверка сим
def sim_reconciliation(driver, path):
    btn_entrance = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Сверка СИМ')]")
    btn_entrance[0].click()
    field = driver.find_element(by=By.NAME, value="filename")
    field.send_keys(path)
    driver.find_element(by=By.XPATH, value="//input[@name='rownumber']").clear()
    row_n = driver.find_elements(by=By.XPATH, value="//input[@name='rownumber']")
    row_n[0].send_keys(1)
    driver.find_element(by=By.XPATH, value="//input[@name='colnumber']").clear()
    col_n = driver.find_elements(by=By.XPATH, value="//input[@name='colnumber']")
    col_n[0].send_keys(2)
    btn_add2 = driver.find_elements(by=By.XPATH, value="//input[@value='Выполнить']")
    btn_add2[0].click()
    while True:
        if check_exists_by_xpath(driver, "//div[@id='errors_txt']"):
            div_err = driver.find_element(by=By.XPATH, value="//div[@id='errors_txt']")
            count = div_err.text.split(": ")[1]
            df = get_exl(path)
            if int(df.shape[0]) == int(count):
                break
        elif check_exists_by_xpath(driver, "//a[contains(text(), 'Скачать файл сверки')]"):
            download_a = driver.find_element(by=By.XPATH, value="//a[contains(text(), 'Скачать файл сверки')]")
            download_a.click()
            file = "/bookv.xls"
            while not os.path.exists(str(Path.home() / "Downloads") + file):
                time.sleep(1)
            os.replace(str(Path.home() / "Downloads") + file, f"{os.getcwd()}/доп{file}")
            del_sim(driver, f"{os.getcwd()}/доп{file}")
            break
        else:
            print("ожидание сверки")
            time.sleep(5)


# удаление
def del_sim(driver, file):
    df = get_exl(file)
    df = df[df[23] == "0"]
    df.to_excel("доп/del_sim.xlsx")
    btn_entrance = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Удаление данных')]")
    btn_entrance[0].click()
    field = driver.find_element(by=By.NAME, value="filename")
    field.send_keys("доп/del_sim.xlsx")
    btn_perform = driver.find_elements(by=By.XPATH, value="//input[@value='Выполнить']")
    btn_perform[0].click()


# Выдача
def issuing_sim(driver, path, mode_find):
    df = get_exl(path)
    try:
        st = False
        # Виртуальные номера
        if df.shape[1] == 6:
            df['Агенты'] = df['Агенты'].str.replace("П/С ", "")
            btn_entrance = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Выдача СИМ')]")
            btn_entrance[0].click()
            excel_num = 0
            agents = df['Агенты'].unique()
            for agent in agents:
                st = False
                new_df = df[df['Агенты'] == agent]
                dates = new_df['Дата'].unique()
                for date in dates:
                    date_df = new_df[new_df['Дата'] == date]
                    date_df.reset_index(inplace=True)
                    for e, number in enumerate(date_df['Номер']):
                        if date_df['SimTrader'][e] != "ст":
                            st = True
                    if st:
                        btn_new_entrance = driver.find_elements(by=By.XPATH, value="//a[@title='Создать элемент']")
                        btn_new_entrance[0].click()

                        if mode_find:
                            try:
                                driver.find_element_by_xpath(
                                    f"//select[@name='contragentid']/option[text()='{agent}']").click()
                            except:
                                return_in_table(driver)
                                break
                        else:
                            inp1 = driver.find_element(by=By.XPATH, value="//input[@id='input1']")
                            inp1.send_keys(agent)

                        inp_date = driver.find_element(by=By.XPATH, value="//input[@name='docdate']")
                        datef = datetime.datetime.strptime(str(date), '%Y-%m-%d %H:%M:%S')
                        inp_date.clear()
                        inp_date.send_keys(datef.strftime('%d.%m.%Y'))

                        # ret_d = sp_check(driver)
                        # select2 = Select(driver.find_element(by=By.XPATH, value="//select[@id='select1']"))
                        # select2.select_by_value(ret_d.get(agent.lstrip().rstrip()))

                        btn_add2 = driver.find_element(by=By.XPATH, value="//input[@value='Записать']")
                        btn_add2.click()

                        driver.find_element(by=By.XPATH, value="//textarea[@name='numberlist']").clear()
                        for e, number in enumerate(date_df['Номер']):
                            if date_df['SimTrader'][e] == "ст":
                                continue

                            number_sim = driver.find_element(by=By.XPATH, value="//textarea[@name='numberlist']")
                            number_sim.send_keys(number + "\n")

                            field = driver.find_element(by=By.NAME, value="filename")
                            field.send_keys(path)
                            df.loc[excel_num, 'SimTrader'] = "ст"
                            excel_num += 1
                            # openExcel.update_spreadsheet(path, df)

                        btn_add = driver.find_elements(by=By.XPATH, value="//input[@value='Добавить']")
                        btn_add[0].click()
                        if check_exists_by_xpath(driver, "//div[@id = 'errors_txt']"):
                            errorNums = re.finditer(r"\(?\b[2-9][0-9]{2}\)?[-. ]?[2-9][0-9]{2}[-. ]?[0-9]{4}\b", driver.find_element(by=By.XPATH, value="//input[@value='Добавить']").text)
                            df[df['номер' in errorNums], 'SimTrader'] = ""

                        btn_add3 = driver.find_element(by=By.XPATH, value="//input[@value='OK']")
                        btn_add3.click()
                        # time.sleep(1)
        # Обычные номера
        elif df.shape[1] == 7:
            df['Агенты'] = df['Агенты'].str.replace("П/С ", "")
            for e, i in enumerate(df['Номер2']):
                if df['SimTrader'][e] == "ст":
                    continue

                btn_entrance = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Выдача СИМ')]")
                btn_entrance[0].click()
                btn_new_entrance = driver.find_elements(by=By.XPATH, value="//a[@title='Создать элемент']")
                btn_new_entrance[0].click()
                driver.find_element_by_xpath(
                    f"//select[@name='contragentid']/option[text()='{df['Агенты'][e]}']").click()

                # if('Модемы. ' in df[9][e]):
                #     inp1 = driver.find_element(by=By.XPATH, value="//input[@id='input1']")
                #     inp1.send_keys(df[9][e])
                #     time.sleep(2)
                # else:
                #     driver.find_element_by_xpath(f"//select[@name='contragentid']/option[text()='{df[9][e]}']").click()
                #     # inp1 = driver.find_element(by=By.XPATH, value="//input[@id='input1']")
                #     # inp1.send_keys(df[9][e])
                #     time.sleep(10)
                # if not ret_d.get(df[9][e].lstrip().rstrip()):
                #     from main import ErrorDialog
                #     er = ErrorDialog("Такого агента не существует: " + df[9][e])
                #     er.exec()
                #     break
                inp_date = driver.find_element(by=By.XPATH, value="//input[@name='docdate']")
                date = datetime.datetime.strptime(str(df['Дата'][e]), '%Y-%m-%d %H:%M:%S')
                inp_date.clear()
                inp_date.send_keys(date.strftime('%d.%m.%Y'))

                btn_add2 = driver.find_element(by=By.XPATH, value="//input[@value='Записать']")
                btn_add2.click()
                # time.sleep(2)
                if str(i) == "nan":
                    driver.find_element(by=By.XPATH, value="//textarea[@name='numberlist']").clear()
                    number_sim = driver.find_element(by=By.XPATH, value="//textarea[@name='numberlist']")
                    number_sim.send_keys(str(df['Номер1'][e]))
                    # time.sleep(2)
                else:
                    driver.find_element(by=By.XPATH, value="//textarea[@name='newsim']").clear()
                    driver.find_element(by=By.XPATH, value="//textarea[@name='newsim2']").clear()
                    min_sim = driver.find_elements(by=By.XPATH, value="//textarea[@name='newsim']")
                    min_sim[0].send_keys(str(df['Номер1'][e]))
                    max_sim = driver.find_elements(by=By.XPATH, value="//textarea[@name='newsim2']")
                    max_sim[0].send_keys(str(df['Номер2'][e]))
                    # time.sleep(2)

                field = driver.find_element(by=By.NAME, value="filename")
                field.send_keys(path)

                btn_add = driver.find_elements(by=By.XPATH, value="//input[@value='Добавить']")
                btn_add[0].click()
                # time.sleep(1)
                # driver.save_screenshot(f"отчет_контрагент{e}.png")

                btn_add3 = driver.find_element(by=By.XPATH, value="//input[@value='OK']")
                btn_add3.click()

                df.loc[e, 'SimTrader'] = "ст"
                # openExcel.update_spreadsheet(path, df)
        else:
            from main import ErrorDialog
            er = ErrorDialog("Неверный формат файла")
            er.exec()
    except Exception as e:
        openExcel.update_spreadsheet(path, df)
        from main import ErrorDialog
        if str(e) == "list index out of range":
            er = ErrorDialog("Неверный логин, пароль или капча")
            er.exec()
        elif "Permission denied" in str(e):
            er = ErrorDialog("Невозможно перезаписать файл, когда он открыт")
            er.exec()
        else:
            er = ErrorDialog(str(e))
            er.exec()
    else:
        openExcel.update_spreadsheet(path, df)


# def sp_check(driver):
#     s1 = (driver.find_elements(by=By.XPATH, value="//select[@id='select1']//option"))
#     ret_d = {}
#     for s in s1:
#         ret_d[s.text.replace("С/П ", "")] = s.get_attribute("value")
#     return ret_d


# Возврат
def sim_refund(driver, path):
    try:
        df = get_exl(path)
        if df.shape[1] == 5:
            df[4] = df[4].str.replace("П/С ", "")
            for e, op in enumerate(df[4]):
                # if df[0][e] == "ст":
                #     continue

                btn_entrance = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Возврат СИМ')]")
                btn_entrance[0].click()
                btn_new_entrance = driver.find_elements(by=By.XPATH, value="//*[contains(text(), 'Новый')]")
                btn_new_entrance[0].click()

                inp1 = driver.find_element(by=By.XPATH, value="//input[@id='input1']")
                inp1.send_keys(df[4][e])
                # if not ret_d.get(df[4][e].lstrip().rstrip()):
                #     from main import ErrorDialog
                #     er = ErrorDialog("Такого агента не существует: " + df[4][e])
                #     er.exec()
                #     break
                # select2 = Select(driver.find_element(by=By.XPATH, value="//select[@id='select1']"))
                # select2.select_by_value(ret_d.get(df[4][e].lstrip().rstrip()))
                btn_add2 = driver.find_element(by=By.XPATH, value="//input[@value='Записать']")
                btn_add2.click()

                driver.find_element(by=By.XPATH, value="//textarea[@name='numberlist']").clear()
                number_sim = driver.find_element(by=By.XPATH, value="//textarea[@name='numberlist']")

                number_sim.send_keys(str(df[3][e]))

                field = driver.find_element(by=By.NAME, value="filename")
                field.send_keys(path)

                btn_add = driver.find_elements(by=By.XPATH, value="//input[@value='Добавить']")
                btn_add[0].click()
                # driver.save_screenshot(f"отчет_контрагент{e}.png")
                btn_add3 = driver.find_element(by=By.XPATH, value="//input[@value='OK']")
                btn_add3.click()

                # df.loc[e, 0] = "ст"
            # df.to_excel(path, index=False, header=None)

        elif df.shape[1] == 10:
            df[9] = df[9].str.replace("П/С ", "")
            for e, i in enumerate(df[7]):
                # if df[0][e] == "ст":
                #     continue

                btn_entrance = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Возврат СИМ')]")
                btn_entrance[0].click()
                btn_new_entrance = driver.find_elements(by=By.XPATH, value="//*[contains(text(), 'Новый')]")
                btn_new_entrance[0].click()

                inp1 = driver.find_element(by=By.XPATH, value="//input[@id='input1']")
                inp1.send_keys(df[9][e])
                # ret_d = sp_check(driver)
                # select2 = Select(driver.find_element(by=By.XPATH, value="//select[@id='select1']"))
                # select2.select_by_value(ret_d.get(df[9][e].lstrip().rstrip()))
                btn_add2 = driver.find_element(by=By.XPATH, value="//input[@value='Записать']")
                btn_add2.click()
                time.sleep(2)
                if str(i) == "nan":
                    driver.find_element(by=By.XPATH, value="//textarea[@name='numberlist']").clear()
                    number_sim = driver.find_element(by=By.XPATH, value="//textarea[@name='numberlist']")
                    number_sim.send_keys(str(df[6][e]))
                else:
                    driver.find_element(by=By.XPATH, value="//textarea[@name='newsim']").clear()
                    driver.find_element(by=By.XPATH, value="//textarea[@name='newsim2']").clear()
                    min_sim = driver.find_elements(by=By.XPATH, value="//textarea[@name='newsim']")
                    min_sim[0].send_keys(str(df[6][e]))
                    max_sim = driver.find_elements(by=By.XPATH, value="//textarea[@name='newsim2']")
                    max_sim[0].send_keys(str(df[7][e]))

                field = driver.find_element(by=By.NAME, value="filename")
                field.send_keys(path)

                btn_add = driver.find_elements(by=By.XPATH, value="//input[@value='Добавить']")
                btn_add[0].click()
                # driver.save_screenshot(f"отчет_контрагент{e}.png")

                btn_add3 = driver.find_element(by=By.XPATH, value="//input[@value='OK']")
                btn_add3.click()

                # df.loc[e, 0] = "ст"
            # df.to_excel(path, index=False, header=None)
        else:
            from main import ErrorDialog
            er = ErrorDialog("Неверный формат файла")
            er.exec()
    except Exception as e:
        from main import ErrorDialog
        if str(e) == "list index out of range":
            er = ErrorDialog("Неверный логин или пароль или капча")
            er.exec()
        else:
            er = ErrorDialog(str(e))
            er.exec()


def check_exists_by_xpath(driver, xpath):
    try:
        driver.find_element(by=By.XPATH, value=xpath)
    except NoSuchElementException:
        return False
    return True


# чтение из xlsx
def get_exl(path):
    return pd.read_excel(path, dtype=str)


# получения ключа по значению
def get_key(d, value):
    for k, v in d.items():
        if v == value:
            return k


def return_in_table(driver):
    driver.find_element(by=By.XPATH, value="//input[@value='OK']").click()

    td = driver.find_elements(by=By.XPATH, value="//a[@'linkcontragentid=0' in ondbClick]//td")
    td[10].find_elements(by=By.XPATH, value='//a[contains(text(), "Удалить")]').click()

    driver.find_element(by=By.XPATH, value="//input[@value = 'Удалить']").click()


if __name__ == '__main__':
    # driver = get_c()
    # cap = input("Cимволы с картинки:")
    # login = input("Логин:")
    # password = input("Пароль:")
    # log_in_to_the_system(driver, login, password, cap)
    # issuing_sim(driver, "C:/Users/Danil/PycharmProjects/Mixu_warehouse/доп/выдача.xlsx")
    # log_in_to_the_system(driver, login, password, cap)
    # path = "C:/Users/Danil/PycharmProjects/Mixu_warehouse/доп/реестр.xlsx"
    # sim_refund(driver, path)
    # print(123)

    s = Service('yandexdriver.exe')
    driver = webdriver.Chrome(service=s)
    url = "https://dss.simtrader.ru/modules/auth.php"
    driver.get(url)
    url = "https://translated.turbopages.org/proxy_u/en-ru.ru.b3adc058-6244144c-19367e5b-74722d776562/https/" \
          "stackoverflow.com/questions/54734087/" \
          "why-does-selenium-webdriver-opens-new-window-for-every-time-run-the-script-and-h"
    driver.get(url)
