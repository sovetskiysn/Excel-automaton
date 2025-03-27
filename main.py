
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

import os
import win32com.client as win32
from traceback import format_exc

from itertools import chain
from functools import partial


from functions import *




# def testVal(new_value, action_type):
#     if action_type == '1' and (not new_value.isdigit() or new_value!= ','):  # Только цифры и запятая при вводе
#         return False
#     return True

def close_app():

    for workbook in ExcelApp.Workbooks:
        workbook.Close(SaveChanges=False)

    ExcelApp.Quit()
    window.destroy()

def flatten_dict_values(d):
    result = []
    for value in d.values():
        if isinstance(value, dict):  # Если вложенный словарь, рекурсивно проходим
            result.extend(flatten_dict_values(value))
        elif isinstance(value, list):  # Если список, добавляем его элементы
            result.extend(value)
    return result


def create_menu(menu, options):
    """ Рекурсивно создает меню для вложенных опций """
    for key, value in options.items() if isinstance(options, dict) else enumerate(options):
        if isinstance(value, (dict, list)):  # Если вложенная структура
            submenu = tk.Menu(menu, tearoff=False)
            menu.add_cascade(label=key, menu=submenu)
            create_menu(submenu, value)  # Рекурсивный вызов
        else:  # Если конечное действие
            menu.add_command(label=value, command= lambda x=value: selected_function(x))




def selected_function(value):

    action_var.set(value)

    window.focus()

    original_checkbox.pack_forget()
    informational_checkbox.pack_forget()
    additional_column_checkbox.pack_forget()
    delete_info_bar_checkbox.pack_forget()
    with_file_name_checkbox.pack_forget()
    columns_entry.pack_forget()

    if value == 'Конвертировать Отчет в Таблицу':
        original_checkbox.pack(side='top', anchor='nw')
        informational_checkbox.pack(side='top', anchor='nw')
        additional_column_checkbox.pack(side='top', anchor='nw')

    elif value == 'Конвертировать Таблицу в Отчет':
        original_checkbox.pack(side='top', anchor='nw')

    elif value == 'Объединить таблицы в одну':
        pass

    elif value == 'Переименовать файлы':
        columns_entry.pack(side='top', anchor='nw')

    elif value == 'Переименовать лист':
        columns_entry.pack(side='top', anchor='nw')

    elif value == 'Просуммировать таблицу':
        original_checkbox.pack(side='top', anchor='nw')
        columns_entry.pack(side='top', anchor='nw')

    elif value == 'Разделить файл на Листы':
        with_file_name_checkbox.pack(side='top', anchor='nw')

    elif value == 'Убрать пустые колонки и строки':
        original_checkbox.pack(side='top', anchor='nw')
        
    elif value == 'Разъединить обьядененные ячейки с заполнением':
        original_checkbox.pack(side='top', anchor='nw')
        delete_info_bar_checkbox.pack(side='top', anchor='nw')

def select_file():
    file_path_list = filedialog.askopenfilenames(filetypes=[("Excel files", ".xlsx .xls")])
    if len(file_path_list) != 0:

        files_paths_string.set(f'Вы выбрали:\n{'\n'.join(file_path_list)}')

def download_file(LIST_OF_WORKBOOKS, file_name_list):
    
    ExcelApp.DisplayAlerts = True

    # тот случай когда выбирается один файл
    if len(LIST_OF_WORKBOOKS) == 1 and len(file_name_list) == 1 and action_var.get() != 'Разделить файл на Листы':


        outcome_file_path = filedialog.asksaveasfilename(initialfile=file_name_list[0], filetypes=[('Excel Files', '*.xlsx'), ('All Files', '*.*')])
        outcome_file_path = os.path.abspath(outcome_file_path)

        LIST_OF_WORKBOOKS[0].SaveAs(outcome_file_path)
        LIST_OF_WORKBOOKS[0].Close(SaveChanges=False)

    else:

        outcome_folder_path = filedialog.askdirectory()

        for workbook, file_name in zip(LIST_OF_WORKBOOKS, file_name_list):

            full_path = os.path.abspath(os.path.join(outcome_folder_path, file_name))
            full_path += '.xlsx'
            # print(full_path)

            workbook.SaveAs(full_path)
            workbook.Close(SaveChanges=False)


    ExcelApp.DisplayAlerts = False
    informational_label.pack_forget()
    download_button.pack_forget()


def process_file():

    for workbook in ExcelApp.Workbooks:
        workbook.Close(SaveChanges=False)

    try:
        file_path_list = files_paths_string.get().split('\n')[1:]

        if files_paths_string.get() == '' : raise Exception('Не выбран путь к файлу')
        if action_var.get() not in flatten_dict_values(action_options): raise Exception('Не выбранно действие')


        LIST_OF_WORKBOOKS = []
        file_name_list = []

        if action_var.get() == 'Объединить файлы в один файл':

            if len(file_path_list) <= 1: raise Exception('Выбранно не достаточно файлов для Объединения')
            
            LIST_OF_WORKBOOKS, file_name_list = combine_files(file_path_list = file_path_list, ExcelApp=ExcelApp)

        else:

            option_and_function = {
                'Конвертировать Отчет в Таблицу': partial(compress_headers, original_sheet=original_var.get(), informational_sheet=informational_var.get(), additional_column=additional_column_var.get()),

                'Конвертировать Таблицу в Отчет': partial(expand_headers, original_sheet=original_var.get()),
                
                'Просуммировать таблицу': partial(groupby_table, columns_string=columns_entry.get(), original_sheet=original_var.get()),

                'Убрать пустые колонки и строки': partial(delete_blank_cols_and_rows, original_sheet=original_var.get()),

                'Разъединить обьядененные ячейки с заполнением': partial(unmerge_the_merged_cells_with_filling, original_sheet=original_var.get(), delete_info_bar=delete_info_bar_var.get()),

                'Разделить файл на файлы': partial(split_file_into_sheets, with_file_name=with_file_name_var.get()),

                'Переименовать файлы': partial(rename_file, name_pattern=columns_entry.get()),

                'Переименовать лист': partial(rename_sheets, name_pattern=columns_entry.get())
            }

            if action_var.get() not in option_and_function.keys():
                raise Exception('Еще не реализованное действие')
            
            unknown_process_function =  option_and_function[action_var.get()]

            for path in file_path_list:

                temp_workbooks_list, temp_names_list = unknown_process_function(file_path = path, ExcelApp=ExcelApp)
                LIST_OF_WORKBOOKS += temp_workbooks_list
                file_name_list += temp_names_list
                # file_name_list.append(os.path.basename(path))


            # print(file_name_list)
            
            # проверка на то если тут одинаковые имена у файлов
            if len(file_name_list) != len(set(file_name_list)):
                raise Exception('Обнаружены листы с одинаковыми названиями')



    except Exception as e:

        if str(e) == 'Не выбран путь к файлу':
            informational_label.configure(text='Сначала выберите файл!!!')

        elif str(e) == 'Не выбранно действие':
            informational_label.configure(text='Выберите действие!!!')

        elif str(e) == 'Невозможно развернуть заголовоки так как обнаруженна панель данных отчета':
            informational_label.configure(text='Прежде чем разворачивать заголовки уберите панель данных отчета')

        elif str(e) == 'Обнаружены листы с одинаковыми названиями':
            informational_label.configure(text='Обнаружены листы с одинаковыми названиями, попробуйте добавить название файла')

        elif str(e) == 'Еще не реализованное действие':
            informational_label.configure(text='Еще не реализованное действие')
        
        elif str(e) == 'Неправильно набранная команда':
            informational_label.configure(text='Неправильно набранная команда')

        else:
            print(format_exc())
            error_string = format_exc()
            informational_label.configure(text=error_string)

        for workbook in ExcelApp.Workbooks:
            workbook.Close(SaveChanges=False)

        informational_label.pack(side='top')

    else:
        informational_label.configure(text='Ваш файл готов')
        informational_label.pack(side='top')
        
        download_button.configure(command= lambda: download_file(LIST_OF_WORKBOOKS, file_name_list=file_name_list))
        download_button.pack(side='top')

    



ExcelApp = win32.DispatchEx("Excel.Application")
ExcelApp.Visible = False
ExcelApp.DisplayAlerts = False

# ExcelApp.EnableEvents = False
# ExcelApp.Interactive = False
# ExcelApp.ScreenUpdating = False 




# -------------------------------------------TKINTER
window = tk.Tk()
window.title('Bitch!!! - released by Sanzhar production corporation')
window.geometry('800x500')
styler = ttk.Style()


window.protocol('WM_DELETE_WINDOW', func=close_app)

upper_frame = tk.Frame(master=window)

# FILE PATH TEXT
files_paths_string = tk.StringVar(window)
file_path_text = tk.Label(master=upper_frame, text='Никакой файл не выбран:\n-', textvariable=files_paths_string, justify='left')
file_path_text.pack(side='top', anchor='nw', padx= 10, pady=10)

# INMORNATIONAL LABEL
informational_label = tk.Label(master=upper_frame, text='')

# DOWNLOAD BUTTON
download_button = tk.Button(master=upper_frame, text='Скачать файл')




# --------------FRAME LOWER PANEL
lower_frame = tk.Frame(master=window, background='lightblue', borderwidth=1, relief='solid', height=80)
lower_frame.pack_propagate(False)

# SELECT BUTTON
select_file_button = tk.Button(master=lower_frame, text="Выбрать файл", command=select_file)
select_file_button.pack(side='left', padx=10)

# ACTION DROPDOWN


action_options = {
    'К каждому файлу': [
        'Разделить файл на файлы',
        'Переименовать файлы',
        'Сохранить CSV как EXCEL'
    ],
    'К листам внутри каждого файла': {
        'К обычным листам': [
            'Переименовать лист',
            'Убрать режим Page',
            'Удалить лист',
            'Сохранить лист',
            'Объединить листы',
            'Убрать пустые колонки и строки',
            'Разъединить объединённые ячейки с заполнением'
        ],
        'К Таблицам': [
            'Конвертировать Отчет в Таблицу',
            'Конвертировать Таблицу в Отчет',
            'Объединить таблицы',
            'Просуммировать таблицу',
            'Транспонирование таблицы'
        ]
    },
    'Сразу несколько файлов': [
        'Объединить файлы в один файл',
        'Разделить файл на файлы',
        'Перенести формулы с файла'
    ]
}






action_var = tk.StringVar(window, value='Выбрать действие')
action_menu = tk.OptionMenu(lower_frame, action_var, value='пустышка', command=selected_function)
action_menu['menu'].delete(0, 'end')
action_menu.configure(width=20, anchor='w', highlightbackground='lightblue')
action_menu.pack(side='left', padx=10)
menu = action_menu['menu']

create_menu(menu, action_options)








# -------------------FRAME PARAMETERS
parameters_frame = tk.Frame(master=lower_frame, background='lightblue')

# -----------------------------------CHECKBOXES
original_var = tk.BooleanVar(value=False)
original_checkbox = tk.Checkbutton(parameters_frame, text='Оригинальный лист',variable=original_var, background='lightblue')

informational_var = tk.BooleanVar(value=True)
informational_checkbox = tk.Checkbutton(parameters_frame, text='Информационный лист',variable=informational_var, background='lightblue')

additional_column_var = tk.BooleanVar(value=False)
additional_column_checkbox = tk.Checkbutton(parameters_frame, text='Название файла как колонка',variable=additional_column_var, background='lightblue')

delete_info_bar_var = tk.BooleanVar(value=True)
delete_info_bar_checkbox = tk.Checkbutton(parameters_frame, text='Удалить информационную панель отчета',variable=additional_column_var, background='lightblue')

with_file_name_var = tk.BooleanVar(value=True)
with_file_name_checkbox = tk.Checkbutton(parameters_frame, text='С названием файла',variable=with_file_name_var, background='lightblue')


# ENTRIES

columns_entry = tk.Entry(parameters_frame, validate="key")

# columns_entry.config(validatecommand=(columns_entry.register(testVal), '%P', '%d'))

parameters_frame.pack(side='left', padx=10)

# PROCESS BUTTON
process_button = tk.Button(master=lower_frame, text='Обработать файл', command=process_file)
process_button.pack(side='right', padx=15)




# FRAMES
upper_frame.pack(fill='both', expand=True)
lower_frame.pack(fill='both', expand=False, side='bottom', ipady=10)

window.mainloop()