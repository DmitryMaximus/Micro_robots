import pandas as pd
from tkinter import *
from tkinter.filedialog import askopenfilename
import os

root = Tk()
root.withdraw()
root.update()
file_path = askopenfilename(parent=root, title='Пожалуйста выберите файл с данными')
root.destroy()
try:
    file = open(file_path, mode="r", encoding="ANSI").read()
except Exception:
    print('Возникла ошибка при чтении файла, проверьте, что кодировка файла = ANSI')
    input()

def find_element(elem, delimiter_pos):
    cols_list = ["\t", ". ИНН ", ". ОГРН ", ". Местонахождение: ", ". Основание: "]
    delimiter = cols_list[delimiter_pos]
    if delimiter in elem:
        elem = elem.split(delimiter)[1]
        index_stop = len(elem)
        for cols in cols_list:
            index = elem.find(cols).__index__()
            index_stop = index if (index > 0 and index < index_stop) else index_stop
        return elem[:index_stop]
    else:
        return ''


df = pd.DataFrame([], columns=["Заголовок", "ФИО", "ИНН", "ОГРН", "МЕСТОНАХОЖДЕНИЕ", "ОСНОВАНИЕ"])
file_lines = list(filter(None,file.splitlines()))

try:
    iterator = 0
    for line in file_lines:
        if line[0] in [str(x) for x in range(0,10)]:
            df.at[iterator, "ФИО"] = line.split('.')[1]
            df.at[iterator, "ИНН"] = find_element(line, 1)
            df.at[iterator, "ОГРН"] = find_element(line, 2)
            df.at[iterator, "МЕСТОНАХОЖДЕНИЕ"] = find_element(line, 3)
            df.at[iterator, "ОСНОВАНИЕ"] = find_element(line, 4)

        else:
            df.at[iterator, "Заголовок"] = line

        iterator += 1
    df.to_excel(os.path.dirname(file_path)+"/Result.xlsx", index=False)
except Exception:
    print('Программа завершилась с ошибкой')
    input()
print('Выполнение програмы завершено')
input()