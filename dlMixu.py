import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
import openExcel
from selenium.common.exceptions import NoSuchElementException
import openpyxl as ox


def log_in(login2, password2, driver):
    if driver == 0:
        s = Service('chromedriver.exe')
        driver = webdriver.Chrome(service=s)
    url = "https://dl.mixu.ru/login"
    driver.get(url)
    # driver.minimize_window()
    e_login = driver.find_elements(by=By.XPATH, value="//input[@id = 'user_login']")
    e_login[0].send_keys(login2)
    e_password = driver.find_elements(by=By.XPATH, value="//input[@id = 'user_password']")
    e_password[0].send_keys(password2)
    e_submit = driver.find_elements(by=By.XPATH, value="//input[@type = 'submit']")
    e_submit[0].click()
    return driver


# поступление
def entrance_lc(driver, path):
    try:
        df = get_exl(path)
        list_virt = []
        list_notvir = []
        if df.shape[1] == 5 or df.shape[1] == 4:
            ind_sim = df[df[0].isnull()].index
            list_virt = df.iloc[ind_sim, 1]

            ind_notvir = df[0].dropna().index
            list_notvir = df.iloc[ind_notvir, 1]
        else:
            from main import ErrorDialog
            er = ErrorDialog("Неверный формат файла")
            er.exec()
            return 0

        if list(list_notvir):
            e_btn1 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Загрузить номера')]")
            e_btn1[0].click()
            time.sleep(1)
            e_btn2 = driver.find_element(by=By.XPATH, value="//span[@id = 'select2-move_to_stock_id-container']")
            e_btn2.click()

            e_btn3 = driver.find_elements(by=By.XPATH, value="//li[contains(text(), 'Склад В2В')]")
            e_btn3[0].click()
            e_btn4 = driver.find_elements(by=By.XPATH, value="//textarea[@id = 'phone_numbers']")
            e_btn4[0].send_keys("\n".join(list(list_notvir)))
            time.sleep(5)
            # e_btn_go = driver.find_elements(by=By.XPATH, value="//input[@data-disable-with = 'Продолжить']")
            # e_btn_go[0].click()
            time.sleep(5)

        if list(list_virt):
            e_btn6 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Загрузить номера')]")
            e_btn6[0].click()
            time.sleep(1)
            e_btn7 = driver.find_element(by=By.XPATH, value="//span[@id = 'select2-move_to_stock_id-container']")
            e_btn7.click()

            e_btn8 = driver.find_elements(by=By.XPATH, value="//li[contains(text(), 'Склад виртуальных номеров Р2')]")
            e_btn8[0].click()
            e_btn9 = driver.find_elements(by=By.XPATH, value="//textarea[@id = 'phone_numbers']")
            e_btn9[0].send_keys("\n".join(list(list_virt)))
            # e_btn_go = driver.find_elements(by=By.XPATH, value="//input[@data-disable-with = 'Продолжить']")
            # e_btn_go[0].click()
            time.sleep(5)
    except Exception as e:
        from main import ErrorDialog
        if str(e) == "list index out of range":
            er = ErrorDialog("Неверный логин или пароль")
            er.exec()
        else:
            er = ErrorDialog(str(e))
            er.exec()

    # e_btn1 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Загрузить номера')]")
    # e_btn1[0].click()
    # time.sleep(1)
    # e_btn2 = driver.find_element(by=By.XPATH, value="//span[@id = 'select2-move_to_stock_id-container']")
    # e_btn2.click()
    # e_btn3 = driver.find_elements(by=By.XPATH, value="//li[contains(text(), 'Склад виртуальных номеров Р2')]")
    # e_btn3[0].click()
    # e_btn4 = driver.find_elements(by=By.XPATH, value="//textarea[@id = 'phone_numbers']")
    # e_btn4[0].send_keys("\n".join(list(list_virt)))


# выдача
def extradition_lc(driver, path):
    df = None
    try:
        df = get_exl(path)
        dfError = []
        # обычные номера
        if df.shape[1] == 7:
            agents = df['Агенты']
            for e, agent in enumerate(agents):
                if (df['DLMixu'][e] == "лк") | (str(df['АгентыЛК'][e]) =="nan"):
                    continue
                e_btn1 = driver.find_elements(by=By.XPATH, value="//input[@id = 'sim_number_from']")
                e_btn1[0].clear()
                e_btn1[0].send_keys(df['Номер1'][e])

                e_btn2 = driver.find_elements(by=By.XPATH, value="//input[@id = 'sim_number_to']")
                e_btn2[0].clear()

                if str(df['Номер2'][e]) == "nan":
                    e_btn2[0].send_keys(df['Номер1'][e])
                else:
                    e_btn2[0].send_keys(df['Номер2'][e])
                e_btn3 = driver.find_elements(by=By.XPATH, value="//input[@value = 'Найти']")
                e_btn3[0].click()

                # Ошибка при отсуствие номеров
                if not check_exists_by_xpath(driver, "//label[@for = 'check_all']"):
                    dfError.append(df['Номер1'][e])
                    continue

                e_btn4 = driver.find_elements(by=By.XPATH, value="//label[@for = 'check_all']")
                e_btn4[0].click()

                e_btn5 = driver.find_elements(by=By.XPATH, value="//input[@value = 'Передать на склад / дилеру']")
                e_btn5[0].click()
                time.sleep(1)

                e_btn6 = driver.find_element(by=By.XPATH, value="//span[@id = 'select2-dealer-container']")
                e_btn6.click()
                e_btn7 = driver.find_elements(by=By.XPATH, value="//input[@class = 'select2-search__field']")
                e_btn7[0].send_keys(str(agent).lstrip().rstrip())

                # Ошибка при отсутсвие агента
                if not check_exists_by_xpath(driver, "//li[@class = 'select2-results__option "
                                                     "select2-results__option--highlighted']"):
                    dfError.append(df['Номер1'][e])
                    e_btn9 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Назад')]")
                    e_btn9[0].click()
                    continue

                e_btn8 = driver.find_elements(by=By.XPATH, value="//li[@class = 'select2-results__option "
                                                                 "select2-results__option--highlighted']")
                e_btn8[0].click()

                e_btn9 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Назад')]")
                e_btn9[0].click()

                # e_btn9 = driver.find_elements(by=By.XPATH, value="//input[@data-disable-with = 'Выполняется...']")
                # e_btn9[0].click()
                # e_btn10 = driver.find_element(by=By.XPATH, value="//a[@class = 'btn modal-action modal-close']")
                # e_btn10.click()

                df.loc[e, 'DLMixu'] = "лк"
                time.sleep(1)

            if len(dfError) > 0:
                from main import ErrorDialog
                er = ErrorDialog(f"{len(dfError)} записей не прошли в систему!")
                er.exec()
                # openExcel.update_spreadsheet(openExcel.copyExFile('\\доп\\выдача.xlsx'), df[df['Номер1'].isin(dfError)])

        # вирутальные номера
        elif df.shape[1] == 6:
            agents = df['АгентыЛК'].dropna().unique()

            for e, agent in enumerate(agents):
                list_sim = df.loc[(df['АгентыЛК'] == agent) & (df['DLMixu'] != "лк"), 'Номер']
                ind_agent = df[(df['АгентыЛК'] == agent) & (df['DLMixu'] != "лк")].index
                if(len(list_sim) == 0):
                    continue
                e_btn0 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Выбрать номера из файла')]")
                e_btn0[0].click()
                time.sleep(0.5)

                e_btn6 = driver.find_elements(by=By.XPATH, value="//textarea[@id = 'phone_numbers']")
                e_btn6[0].send_keys("\n".join(list(list_sim)))

                e_btn7 = driver.find_elements(by=By.XPATH, value="//input[@value = 'Продолжить']")
                e_btn7[0].click()
                time.sleep(count_delit(len(list_sim), 1))

                if check_exists_by_xpath(driver, "//div[@class = 'row not_found_phones']"):
                    list_number = []
                    er_number = driver.find_elements(by=By.XPATH, value="//div[@class = 'row not_found_phones']//ul//li")
                    for num in er_number:
                        list_number.append(num.text)
                    dfError = dfError + list_number

                    e_btn9 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Назад')]")
                    e_btn9[0].click()
                    # from main import ErrorDialog
                    # er = ErrorDialog(f"Сим-карты не найдены: {list_number}")
                    # er.exec()
                    continue
                time.sleep(0.5)
                e_btn6 = driver.find_element(by=By.XPATH, value="//span[@id = 'select2-dealer-container']")
                e_btn6.click()

                e_btn7 = driver.find_elements(by=By.XPATH, value="//input[@class = 'select2-search__field']")
                e_btn7[0].send_keys(agent.lstrip().rstrip())
                # Ошибка Не найден Агент

                if check_exists_by_xpath(driver, "//li[@class = 'select2-results__option select2-results__message']"):
                    dfError = dfError + list(df.loc[df['АгентыЛК'] == agent, 'Номер'])

                    # from main import ErrorDialog
                    # er = ErrorDialog(f"Агент {agent.lstrip().rstrip()} не найден")
                    # er.exec()
                    time.sleep(1)
                    e_btn9 = driver.find_element(by=By.XPATH, value="//a[contains(text(), 'Назад')]")
                    e_btn9.click()
                    continue
                e_btn8 = driver.find_elements(by=By.XPATH, value="//li[@class = 'select2-results__option select2-results__option--highlighted']")
                e_btn8[0].click()

                e_btn9 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Назад')]")
                e_btn9[0].click()


                # e_btn9 = driver.find_elements(by=By.XPATH, value="//input[@data-disable-with = 'Выполняется...']")
                # e_btn9[0].click()
                # e_btn10 = driver.find_element(by=By.XPATH, value="//a[@class = 'btn modal-action modal-close']")
                # e_btn10.click()

                df.loc[ind_agent, 'DLMixu'] = "лк"
                # openExcel.update_spreadsheet(path, df)
            if len(dfError) > 0:
                from main import ErrorDialog
                er = ErrorDialog(f"{len(dfError)} записей не прошли в систему!\n")
                er.exec()
                print(list(dfError))
                df['АгентыЛК'] = df['Агенты']
                openExcel.update_spreadsheet(openExcel.copyExFile('\\доп\\вирт_сим.xlsx'), df[df['Номер'].isin(dfError)])

        else:
            from main import ErrorDialog
            er = ErrorDialog("Неверный формат файла")
            er.exec()
    except Exception as e:
        openExcel.update_spreadsheet(path, df)
        from main import ErrorDialog
        if str(e) == "list index out of range":
            er = ErrorDialog("Неверный логин или пароль")
            print(str(e))
            er.exec()
        else:
            er = ErrorDialog(str(e))
            print(str(e))
            er.exec()
    else:
        openExcel.update_spreadsheet(path, df)



# возврат
def refund_lc(driver, path):
    try:
        df = get_exl(path)
        # обычные
        if df.shape[1] == 10:
            agents = df[9]
            for e, agent in enumerate(agents):
                e_btn1 = driver.find_elements(by=By.XPATH, value="//input[@id = 'sim_number_from']")
                e_btn1[0].clear()
                e_btn1[0].send_keys(df[6][e])

                e_btn2 = driver.find_elements(by=By.XPATH, value="//input[@id = 'sim_number_to']")
                e_btn2[0].clear()

                if str(df[7][e]) == "nan":
                    e_btn2[0].send_keys(df[6][e])
                else:
                    e_btn2[0].send_keys(df[6][e])

                e_btn3 = driver.find_elements(by=By.XPATH, value="//input[@value = 'Найти']")
                e_btn3[0].click()

                if not check_exists_by_xpath(driver, "//label[@for = 'check_all']"):
                    from main import ErrorDialog
                    er = ErrorDialog(f"Сим-карты не найдены")
                    er.exec()
                    break
                e_btn4 = driver.find_elements(by=By.XPATH, value="//label[@for = 'check_all']")
                e_btn4[0].click()

                e_btn5 = driver.find_elements(by=By.XPATH, value="//input[@value = 'Передать на склад / дилеру']")
                e_btn5[0].click()
                time.sleep(2)

                e_btn6 = driver.find_element(by=By.XPATH, value="//span[@id = 'select2-stock-container']")
                e_btn6.click()

                e_btn7 = driver.find_elements(by=By.XPATH, value="//input[@class = 'select2-search__field']")
                e_btn7[0].send_keys("Склад В2В")

                e_btn8 = driver.find_elements(by=By.XPATH, value="//li[@class = 'select2-results__option select2-results__option--highlighted']")
                e_btn8[0].click()
                time.sleep(2)
                # e_btn9 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Назад')]")
                # e_btn9[0].click()
                e_btn9 = driver.find_elements(by=By.XPATH, value="//input[@data-disable-with = 'Выполняется...']")
                e_btn9[0].click()
        # вирт
        elif df.shape[1] == 5:
            agents = df[4].unique()
            for agent in agents:
                list_sim = df[df[4] == agent][3]

                e_btn0 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Выбрать номера из файла')]")
                e_btn0[0].click()
                time.sleep(2)

                e_btn6 = driver.find_elements(by=By.XPATH, value="//textarea[@id = 'phone_numbers']")
                e_btn6[0].send_keys("\n".join(list(list_sim)))

                e_btn7 = driver.find_elements(by=By.XPATH, value="//input[@value = 'Продолжить']")
                e_btn7[0].click()
                time.sleep(2)
                if check_exists_by_xpath(driver, "//div[@class = 'row not_found_phones']"):
                    list_number = []
                    er_number = driver.find_elements(by=By.XPATH, value="//div[@class = 'row not_found_phones']//ul//li")
                    for num in er_number:
                        list_number.append(num.text)
                    from main import ErrorDialog
                    er = ErrorDialog(f"Номера не найдены: {list_number}")
                    er.exec()
                    break

                e_btn6 = driver.find_element(by=By.XPATH, value="//span[@id = 'select2-stock-container']")
                e_btn6.click()

                e_btn7 = driver.find_elements(by=By.XPATH, value="//input[@class = 'select2-search__field']")
                e_btn7[0].send_keys("Склад виртуальных номеров Р2")

                e_btn8 = driver.find_elements(by=By.XPATH, value="//li[@class = 'select2-results__option "
                                                                 "select2-results__option--highlighted']")
                e_btn8[0].click()

                # e_btn9 = driver.find_elements(by=By.XPATH, value="//a[contains(text(), 'Назад')]")
                # e_btn9[0].click()
                e_btn9 = driver.find_elements(by=By.XPATH, value="//input[@data-disable-with = 'Выполняется...']")
                e_btn9[0].click()
        else:
            from main import ErrorDialog
            er = ErrorDialog("Неверный формат файла")
            er.exec()
    except Exception as e:
        from main import ErrorDialog
        if str(e) == "list index out of range":
            er = ErrorDialog("Неверный логин или пароль")
            er.exec()
        else:
            er = ErrorDialog(str(e))
            er.exec()


def get_exl(path):
    return pd.read_excel(path, dtype=str)


def check_exists_by_xpath(driver, xpath):
    try:
        driver.find_element(by=By.XPATH, value=xpath)
    except NoSuchElementException:
        return False
    return True


def count_delit(number, count):
    while number > 20:
        number -= 20
        count += 1
    return count


if __name__ == "__main__":
    login = "test_2022"
    password = "test_2022"
    # entrance_lc(driver=log_in(login, password))
    refund_lc(driver=log_in(login, password), path="доп/выдача.xlsx")

