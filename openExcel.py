import openpyxl as ox
import pandas as pd
import os
from tkinter import filedialog
from tkinter import *
import shutil


def copyExFile(fileName):
    root = Tk()
    files = [('Таблицы Excel', '*.xlsx')]

    root.filename = filedialog.asksaveasfile(initialdir=os.getcwd(), filetypes=files, title="Сохранить", defaultextension=files)
    # df = pd.read_excel(os.getcwd() + path)
    # df.drop(df.index, inplace=True)
    shutil.copy(os.getcwd() + fileName, root.filename.name)
    root.destroy()
    return root.filename.name


def update_spreadsheet(path: str, _df, starcol: int = 1, startrow: int = 2, sheet_name: str = "Sheet1"):
    '''

    :param path: Путь до файла Excel
    :param _df: Датафрейм Pandas для записи
    :param starcol: Стартовая колонка в таблице листа Excel, куда буду писать данные
    :param startrow: Стартовая строка в таблице листа Excel, куда буду писать данные
    :param sheet_name: Имя листа в таблице Excel, куда буду писать данные
    :return:
    '''
    wb = ox.load_workbook(path)
    for ir in range(0, len(_df)):
        for ic in range(0, len(_df.iloc[ir])):
            wb[sheet_name].cell(startrow + ir, starcol + ic).value = _df.iloc[ir][ic]
    wb.save(path)


