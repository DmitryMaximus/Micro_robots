import pandas as pd
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.simpledialog import askinteger
import os

root = Tk()
root.withdraw()
root.update()
file_path = askopenfilename(parent=root, title='Пожалуйста выберите файл с данными')
cols_count = askinteger(parent=root,title='Ввод данных', prompt='Введите количество столбцов в таблице')
root.destroy()
file = open(file_path, mode="r", encoding="utf-8").read()

try:
    df = pd.DataFrame([], columns=[str(i) for i in range(1, cols_count + 1)])
    iterator = 0
    for line in file.splitlines():
        element = line.split('|')
        element=element[1:]
        element=element[:-1]
        if element[0].replace(" ",'')!='':
            for elem, col in zip(element,list(df)):
                df.at[iterator,col] = elem
            iterator += 1
        else:
            for elem, col in zip(element,list(df)):
                df.at[iterator-1,col]=df[col].iloc[iterator-1]+elem
    df.to_excel(os.path.dirname(file_path)+"/Result.xlsx", index=False)
except Exception:
    print('Программа завершилась с ошибкой')
    input()
print('Выполнение програмы завершено')
input()