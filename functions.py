
import os
import time

import pandas as pd


def inner_clean_sheet(Sheet, ExcelApp):

    # unhide
    Sheet.Columns.EntireColumn.Hidden = False
    Sheet.Rows.EntireColumn.Hidden = False

    # find rows and column which have empties, ignore that doesn't. it need because of optimization
    try:
        only_blank_cells = Sheet.UsedRange.Cells.SpecialCells(4) # 4 means blank cells.
    
    except Exception as e:
        pass
    else:
        removal_col = set()
        removal_row = set()

        # only_blank_cells - multiple range
        for i_range in only_blank_cells.Areas:
            removal_col.update([i_range.Columns(col_i).EntireColumn.Column for col_i in range(1, i_range.Columns.Count + 1)])
            removal_row.update([i_range.Rows(row_i).EntireRow.Row for row_i in range(1, i_range.Rows.Count + 1)])


        # Check if needs to remove cols and remove appreciate cols
        if removal_col:
            removal_col = sorted([index for index in removal_col if ExcelApp.WorksheetFunction.CountA(Sheet.UsedRange.Columns(index).EntireColumn) == 0], reverse=True)
        if removal_col:
            for i in removal_col:
                Sheet.UsedRange.Columns(i).EntireColumn.Delete()

        # Check if needs to remove cols and remove appreciate cols
        if removal_row:
            removal_row = sorted([index for index in removal_row if ExcelApp.WorksheetFunction.CountA(Sheet.UsedRange.Rows(index).EntireRow) == 0], reverse=True)
        if removal_row:
            for i in removal_row:
                Sheet.UsedRange.Rows(i).EntireRow.Delete()


    # start_time = time.monotonic()
    # for i in range(1, Sheet.UsedRange.Columns.Count + 1):


    #     if Sheet.UsedRange.Columns(i).ColumnWidth <= 3.0:
    #         Sheet.UsedRange.Columns(i).ColumnWidth = 3.0
    # end_time = time.monotonic()
    # print(f'выравнивание   -   {round(end_time-start_time,3)} \n')

def inner_unmerge_with_filling(Sheet):
    # unmerge with filling all cells
    only_blank_cells = Sheet.UsedRange.Cells.SpecialCells(4)

    for cell in only_blank_cells:
        if cell.MergeCells:
            temp_text = cell.MergeArea.Cells(1).Value
            merged_area = cell.MergeArea
            cell.UnMerge()

            for unmerged_cell in merged_area.Cells:
                unmerged_cell.value = temp_text

def inner_style_headers(Range, Alignment=True, AutoFit = True, WrapText = True):

    if Alignment == True:
        # Alignment of text.  -4108 means center
        Range.VerticalAlignment = -4108
        Range.HorizontalAlignment = -4108

    if WrapText == True:
        Range.WrapText = True

    if AutoFit == True:
        Range.Worksheet.UsedRange.Columns.AutoFit() 
        Range.Worksheet.UsedRange.Rows.AutoFit()
        

    # 1 means typical border
    Range.Worksheet.UsedRange.Borders.LineStyle = 1

def inner_remove_infobar_if_exists(Sheet, file_path):

    parameters = ['Время выдачи', 'Дата формирования отчета']
    finds = [Sheet.UsedRange.Find(What=param, LookAt=2) for param in parameters]
    infobar_end_cell = next((i for i in finds if i), None)

    infobar_list = [f'Имя исходного файла: {os.path.basename(file_path)}']

    if infobar_end_cell:
        for i in range(infobar_end_cell.Row):

            first_cell = Sheet.Cells(1,1)
            if first_cell.Value != '' and first_cell.Value != None:
                infobar_list.append(first_cell.Value)
            first_cell.EntireRow.Delete()

        infobar_list[1] = f'Название: {infobar_list[1].replace('\n','')}'

        for index, elem in enumerate(infobar_list[2:], start=2):
            infobar_list[index: index+1] = elem.split('\n')


    return infobar_list

def inner_copy_file(file_path, ExcelApp):

    # открываем файлы
    WB_original = ExcelApp.Workbooks.Open(file_path)
    WB_working = ExcelApp.Workbooks.Add()

    # копируем sheets
    for sheet in WB_original.Worksheets:
        sheet.Copy(WB_working.Worksheets(WB_working.Worksheets.Count))

    WB_original.Close(SaveChanges=False)
    WB_working.Worksheets(WB_working.Worksheets.Count).Delete()

    return WB_working

# --------------------------------------------------------------------------------------------------------------------------------------------------------


# -----К каждому файлу

def rename_file(file_path, ExcelApp, name_pattern):
    
    # открываем файлы
    WB_original = ExcelApp.Workbooks.Open(file_path)
    WB_working = ExcelApp.Workbooks.Add()

    # копируем sheets
    for sheet in WB_original.Worksheets:
        sheet.Copy(WB_working.Worksheets(WB_working.Worksheets.Count))

    WB_original.Close(SaveChanges=False)
    WB_working.Worksheets(WB_working.Worksheets.Count).Delete()
    
    # WB_original   |     WB_working
    OLD_NAME = os.path.splitext(os.path.basename(file_path))[0]
    new_name = eval(name_pattern)

    if not isinstance(new_name, str): raise Exception('Неправильно набранная команда')

    new_name = new_name.replace("<", "").replace(">", "").replace(":", "").replace('"', "").replace("/", "").replace("\\", "").replace("|", "").replace("?", "").replace("*", "").rstrip(". ")


    
    return [WB_working], [new_name]

# -----------К листам внутри каждого файла



def compress_headers(file_path, ExcelApp, original_sheet=True, informational_sheet=True, additional_column=True):

    # открываем файлы
    WB_original = ExcelApp.Workbooks.Open(file_path)
    WB_working = ExcelApp.Workbooks.Add()

    # копируем отчет для работы
    WB_original.Worksheets(1).Copy(WB_working.Worksheets(1))
    WB_original.Worksheets(1).Copy(WB_working.Worksheets(1))
    WB_original.Close(SaveChanges=False)

    for index, ws in enumerate(WB_working.Worksheets):
        ws.name = str(index + 1)

    # 31 - 8   | 31- 9  | 31-14
    WB_working.Worksheets(1).name = 'working'
    WB_working.Worksheets(2).name = 'original'
    WB_working.Worksheets(3).name = 'informational'

    WS_working = WB_working.Worksheets(1)
    WS_original = WB_working.Worksheets(2)
    WS_informational = WB_working.Worksheets(3)




    # 1 - preserve and delete infor bar 
    infobar_list = inner_remove_infobar_if_exists(WS_working, file_path)

    # В infobar_list всегда есть как минимум название старого файла
    for index, value in enumerate(infobar_list, 1):
        WS_informational.Cells(index, 1).Value = value
    


    # 2 - clean
    inner_clean_sheet(WS_working, ExcelApp)



            
    # 3 - Header range creating

    # removal_row = WS_working.UsedRange.Find(What='A', LookAt=1) # another needless row
    # if removal_row:
    #     removal_row.EntireRow.Delete()


    first_row_without_merged_cells = None
    for row in WS_working.UsedRange.Rows:

        row_without_merged_cells_flag = True
        for cell in row.Cells:
            if cell.MergeCells:
                row_without_merged_cells_flag = False
                break
        if row_without_merged_cells_flag == True:
            first_row_without_merged_cells = row.Row
            break

    headers_range = WS_working.UsedRange.Rows(f'1:{first_row_without_merged_cells - 1}')
    number_of_level = headers_range.Rows.Count


    # 4 - unmerge
    inner_unmerge_with_filling(WS_working)


    # 5 - Replace previous header with the new
    last_row_number = headers_range.Rows(number_of_level).Row

    WS_working.Rows(last_row_number).GetOffset(1, 0).Insert() 

    values_list = []
    for i in range(1, headers_range.Columns.Count + 1):

        values_in_column = [cell.Value or '' for cell in headers_range.Columns(i).Cells]
        values_list.append(' | '.join(values_in_column).replace('\n', ' '))

    WS_working.UsedRange.Rows(last_row_number).GetOffset(1, 0).Value = values_list # Fill with new header names
    headers_range.EntireRow.Delete() # remove previous header

    headers_range = WS_working.UsedRange.Rows(1)
    last_row_number = headers_range.Rows(headers_range.Rows.Count).Row



    if additional_column == True:

        new_column_range = WS_working.UsedRange.Columns(headers_range.Columns.Count).GetOffset(0, 1)

        headers_range = headers_range.GetResize(RowSize = headers_range.Rows.Count, ColumnSize = headers_range.Columns.Count + 1)

        new_column_range.value = os.path.basename(file_path).split('.')[0]
        new_column_range.Cells(1,1).value = ' | '.join(['Новая колонка' for i in range(number_of_level)])

        WS_working.UsedRange.Cells(last_row_number + 1,1).Copy()

        
        new_column_range.PasteSpecial(-4122)
        # new_column_range.NumberFormat = "General"
        ExcelApp.CutCopyMode = False


        headers_range.Cells(1,1).Copy()
        headers_range.PasteSpecial(-4122)
        ExcelApp.CutCopyMode = False

        new_column_range.Columns(1).Cut()

        WS_working.UsedRange.Columns(1).Insert(Shift= -4161)



    
    if original_sheet == False:
        WS_original.Delete()
        
    if informational_sheet == False:
        WS_informational.Delete()

    inner_style_headers(headers_range)
    

    file_name = os.path.splitext(os.path.basename(file_path))[0]
    return [WB_working], [file_name]

def expand_headers(file_path, ExcelApp, original_sheet=True):
    
    # открываем файлы
    WB_original = ExcelApp.Workbooks.Open(file_path)
    WB_working = ExcelApp.Workbooks.Add()

    # копируем отчет для работы
    WB_original.Worksheets(1).Copy(WB_working.Worksheets(1))
    WB_original.Worksheets(1).Copy(WB_working.Worksheets(1))
    WB_original.Close(SaveChanges=False)


    for index, ws in enumerate(WB_working.Worksheets):
        ws.name = str(index + 1)

    WB_working.Worksheets(1).name = 'working'
    WB_working.Worksheets(2).name = 'original'
    WB_working.Worksheets(3).Delete()

    WS_working = WB_working.Worksheets(1)
    WS_original = WB_working.Worksheets(2)

    # 1 - clean
    inner_clean_sheet(WS_working, ExcelApp)

    # 2 - check if infobar exists
    infobar_list = inner_remove_infobar_if_exists(WS_working, file_path)
    if len(infobar_list) > 1: raise Exception('Невозможно развернуть заголовоки так как обнаруженна панель данных отчета')
        
    # 3 - get values from previous header
    header_row = WS_working.UsedRange.Rows(1)
    column_values_list = [header_row.Columns(i).Value.replace('\n', ' ').split(' | ') for i in range(1, header_row.Columns.Count + 1)]
    number_of_level = max([len(i) for i in column_values_list])

    # 4 - creating new header_range
    WS_working.UsedRange.Rows(f'1:{number_of_level}').EntireRow.Insert()
    header_range = WS_working.UsedRange.Rows(f'1:{number_of_level}').GetOffset(-number_of_level,0) # GetOffset - нужен потому что новые строчки не считаются в UsedRange

    # 5 - transfer values and format from previous header to new one
    for i_col, col in enumerate(header_range.Columns): # заполняем новые строки значениями из column_values_list
        for i_cell,cell in enumerate(col.Cells):

            if column_values_list[i_col][i_cell] != '':
                cell.value = column_values_list[i_col][i_cell]


    header_row.Cells(1, 1).Copy() # copy style
    header_range.PasteSpecial(-4122)
    # header_range.NumberFormat = "General"
    ExcelApp.CutCopyMode = False

    header_row.Delete()

    

    # 6 - merge specific cells
    for cell in header_range.Cells:
        
        if not cell.MergeCells and cell.value != '' and cell.value != None:

            horizontal_i = 0
            while cell.value == cell.GetOffset(0, horizontal_i + 1).value:
                horizontal_i+=1

            vertical_i = 0
            while cell.value == cell.GetOffset(vertical_i + 1,0).value:
                vertical_i+=1

            header_range.Range(cell, cell.GetOffset(vertical_i,horizontal_i)).Merge()



    if original_sheet == False:
        WS_original.Delete()


    inner_style_headers(header_range)
    
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    return [WB_working], [file_name]

def delete_blank_cols_and_rows(file_path, ExcelApp, original_sheet=True):

    # открываем файлы
    WB_original = ExcelApp.Workbooks.Open(file_path)
    WB_working = ExcelApp.Workbooks.Add()

    # копируем отчет для работы
    WB_original.Worksheets(1).Copy(WB_working.Worksheets(1))
    WB_original.Worksheets(1).Copy(WB_working.Worksheets(1))
    WB_original.Close(SaveChanges=False)

    for index, ws in enumerate(WB_working.Worksheets):
        ws.name = str(index + 1)

    WB_working.Worksheets(1).name = 'working'
    WB_working.Worksheets(2).name = 'original'

    WB_working.Worksheets(3).Delete()

    WS_working = WB_working.Worksheets(1)
    WS_original = WB_working.Worksheets(2)


    inner_clean_sheet(WS_working, ExcelApp)

    WS_working.UsedRange.Columns.AutoFit() 
    WS_working.UsedRange.Rows.AutoFit()

    if original_sheet == False:
        WS_original.Delete()

        
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    return [WB_working], [file_name]

def unmerge_the_merged_cells_with_filling(file_path, ExcelApp, original_sheet=True, delete_info_bar=True):

    # открываем файлы
    WB_original = ExcelApp.Workbooks.Open(file_path)
    WB_working = ExcelApp.Workbooks.Add()

    # копируем отчет для работы
    WB_original.Worksheets(1).Copy(WB_working.Worksheets(1))
    WB_original.Worksheets(1).Copy(WB_working.Worksheets(1))
    WB_original.Close(SaveChanges=False)

    for index, ws in enumerate(WB_working.Worksheets):
        ws.name = str(index + 1)

    WB_working.Worksheets(1).name = 'working'
    WB_working.Worksheets(2).name = 'original'
    WB_working.Worksheets(3).Delete()

    WS_working = WB_working.Worksheets(1)
    WS_original = WB_working.Worksheets(2)


    if delete_info_bar == True:
        inner_remove_infobar_if_exists(WS_working, file_path)

    inner_unmerge_with_filling(WS_working)

    if original_sheet == False:
        WS_original.Delete()

        
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    return [WB_working], [file_name]

def groupby_table(file_path, ExcelApp, columns_string, original_sheet=True):

    df = pd.read_excel(file_path, sheet_name=0)
    column_list = list(df.columns.to_numpy())

    column_indexes_list = [int(i)-1 for i in columns_string.split(',') if i!='']

    if len(column_indexes_list) == 1:
        column_indexes_list = [i for i in range(column_indexes_list[0])]


    
    if len(column_list) < max(column_indexes_list): raise Exception('Невозможно развернуть заголовоки так как обнаруженна панель данных отчета')

    column_indexes_list = list(df.columns[column_indexes_list].to_numpy())
    df_my = df.groupby(column_indexes_list, as_index=False, dropna=False).apply(lambda x: x.sum(numeric_only=True))
    df_my = pd.concat([pd.DataFrame([df_my.columns], columns = df_my.columns), df_my], axis=0, ignore_index=True)

    df_my = df_my.fillna('')


    

    # открываем файлы
    WB_original = ExcelApp.Workbooks.Open(file_path)
    WB_working = ExcelApp.Workbooks.Add()

    # копируем отчет для работы
    WB_original.Worksheets(1).Copy(WB_working.Worksheets(1))
    WB_original.Close(SaveChanges=False)

    for index, ws in enumerate(WB_working.Worksheets):
        ws.name = str(index + 1)

    WB_working.Worksheets.Add()


    WB_working.Worksheets(1).name = 'working'
    WB_working.Worksheets(2).name = 'original'
    WB_working.Worksheets(3).Delete()

    WS_working = WB_working.Worksheets(1)
    WS_original = WB_working.Worksheets(2)


    # переносим значения
    StartRow = 1
    StartCol = 1
    WS_working.Range(WS_working.Cells(StartRow,StartCol),# Cell to start the "paste"
            WS_working.Cells(StartRow+len(df_my.index)-1,
                    StartCol+len(df_my.columns)-1)
            ).Value = df_my.values


    # Copy styling from excels


    WS_original.Cells(5,1).Copy()
    WS_working.UsedRange.PasteSpecial(-4122)
    ExcelApp.CutCopyMode = False

    WS_original.Cells(1,1).Copy()
    WS_working.UsedRange.Rows(1).PasteSpecial(-4122)
    ExcelApp.CutCopyMode = False


    WS_working.UsedRange.WrapText = False
    # WS_working.UsedRange.NumberFormat = "General"


    # WS_working.UsedRange.Columns.AutoFit() 
    WS_working.UsedRange.Rows.AutoFit()
    # inner_style_headers(WS_working.UsedRange.Rows(1), WrapText=False)
    

    if original_sheet == False:
        WS_original.Delete()

    file_name = os.path.splitext(os.path.basename(file_path))[0]
    return [WB_working], [file_name]

def rename_sheets(file_path, ExcelApp, name_pattern):

    WB_working = inner_copy_file(file_path=file_path, ExcelApp=ExcelApp)

    FILE_NAME = os.path.splitext(os.path.basename(file_path))[0]

    for sheet in WB_working.Worksheets:

        OLD_NAME = sheet.name
        new_name = eval(name_pattern)

        sheet.name = new_name


    if not isinstance(new_name, str): raise Exception('Неправильно набранная команда')

    
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    return [WB_working], [file_name]




# --------К Таблицам
def combine_table_into_one_through_files(file_path_list, ExcelApp):

    WB_working = ExcelApp.Workbooks.Add()
    WB_working.Worksheets(1).name = 'working'
    WS_working = WB_working.Worksheets(1)

    pd_frames = []

    for path in file_path_list:
        pd_frames.append(pd.read_excel(path, sheet_name=0, dtype=str))

    combined_df = pd.concat(pd_frames, axis=0, ignore_index=True)
    combined_df = pd.concat([pd.DataFrame([combined_df.columns], columns = combined_df.columns), combined_df], axis=0, ignore_index=True)
    combined_df = combined_df.fillna('')
    StartRow = 1
    StartCol = 1


    WS_working.Range(WS_working.Cells(StartRow,StartCol),# Cell to start the "paste"
            WS_working.Cells(StartRow+len(combined_df.index)-1,
                    StartCol+len(combined_df.columns)-1)
            ).Value = combined_df.values
    


    # Copy styling from excels

    WB_first = ExcelApp.Workbooks.Open(file_path_list[0])
    WS_first = WB_first.Worksheets(1)

    WS_first.Cells(5,1).Copy()
    WS_working.UsedRange.PasteSpecial(-4122)
    ExcelApp.CutCopyMode = False

    WS_first.Cells(1,1).Copy()
    WS_working.UsedRange.Rows(1).PasteSpecial(-4122)
    ExcelApp.CutCopyMode = False

    WB_first.Close(SaveChanges=False)

    WS_working.UsedRange.WrapText = False
    # WS_working.UsedRange.NumberFormat = "General"


    # WS_working.UsedRange.Columns.AutoFit() 
    WS_working.UsedRange.Rows.AutoFit()
    # inner_style_headers(WS_working.UsedRange.Rows(1), WrapText=False)
    
    FILE_NAME = ' + '.join([os.path.splitext(os.path.basename(path))[0] for path in file_path_list])

    return WB_working, FILE_NAME


# --------------------------------------------------------------------------------------------------------------------------------------------------------
# Сразу несколько файлов
def combine_files(file_path_list, ExcelApp):

    WB_working = ExcelApp.Workbooks.Add()
    WB_working.Worksheets(1).name = 'working'


    for path in file_path_list:
        WB_original = ExcelApp.Workbooks.Open(path)

        for sheet in WB_original.Worksheets:
            if sheet.name in [WB_working.Worksheets(i).Name for i in range(1, WB_working.Worksheets.Count+1)]:
                raise Exception('Обнаружены листы с одинаковыми названиями')
            else:
                sheet.Copy(WB_working.Worksheets.Count)

        WB_original.Close(SaveChanges=False)


    FILE_NAME = ' + '.join([os.path.splitext(os.path.basename(path))[0] for path in file_path_list])

    return [WB_working], [FILE_NAME]


def split_file_into_sheets(file_path, ExcelApp, with_file_name=True):


    WB_original = ExcelApp.Workbooks.Open(file_path)

    file_name, file_extension = os.path.splitext(os.path.basename(file_path))

    WORKBOOKS_LIST = []
    NAMES_LIST = []

    for i in range(1, WB_original.Worksheets.Count + 1):

        WB_original.Worksheets(i).Visible = True
        WB_original.Worksheets(i).Columns.EntireColumn.Hidden = False
        WB_original.Worksheets(i).Rows.EntireColumn.Hidden = False

        

        worksheet_name = WB_original.Worksheets(i).name
        print(worksheet_name)

        new_workbook = ExcelApp.Workbooks.Add()

        WB_original.Worksheets(i).Copy(new_workbook.Worksheets(1))

        new_workbook.Worksheets(2).Delete()


        WORKBOOKS_LIST.append(new_workbook)

        if with_file_name == True:
            NAMES_LIST.append(file_name + ' - ' + worksheet_name + file_extension)
        else:
            NAMES_LIST.append(worksheet_name)

    WB_original.Close(SaveChanges=False)


    return WORKBOOKS_LIST, NAMES_LIST











def compress_headers_testing(file_path, ExcelApp, original_sheet=True, informational_sheet=True, additional_column=True):


    def compress_headers_on_one_sheet():
        
        print(informational_sheet)


    WB_working = inner_copy_file(file_path=file_path, ExcelApp=ExcelApp)


    compress_headers_on_one_sheet()






    # 31 - 8   | 31- 9  | 31-14
    WB_working.Worksheets(1).name = 'processed'
    WB_working.Worksheets(2).name = 'original'
    WB_working.Worksheets(3).name = 'informational'

    WS_working = WB_working.Worksheets(1)
    WS_original = WB_working.Worksheets(2)
    WS_informational = WB_working.Worksheets(3)




    # 1 - preserve and delete infor bar 
    infobar_list = inner_remove_infobar_if_exists(WS_working, file_path)

    # В infobar_list всегда есть как минимум название старого файла
    for index, value in enumerate(infobar_list, 1):
        WS_informational.Cells(index, 1).Value = value
    


    # 2 - clean
    inner_clean_sheet(WS_working, ExcelApp)



            
    # 3 - Header range creating

    # removal_row = WS_working.UsedRange.Find(What='A', LookAt=1) # another needless row
    # if removal_row:
    #     removal_row.EntireRow.Delete()


    first_row_without_merged_cells = None
    for row in WS_working.UsedRange.Rows:

        row_without_merged_cells_flag = True
        for cell in row.Cells:
            if cell.MergeCells:
                row_without_merged_cells_flag = False
                break
        if row_without_merged_cells_flag == True:
            first_row_without_merged_cells = row.Row
            break

    headers_range = WS_working.UsedRange.Rows(f'1:{first_row_without_merged_cells - 1}')
    number_of_level = headers_range.Rows.Count


    # 4 - unmerge
    inner_unmerge_with_filling(WS_working)


    # 5 - Replace previous header with the new
    last_row_number = headers_range.Rows(number_of_level).Row

    WS_working.Rows(last_row_number).GetOffset(1, 0).Insert() 

    values_list = []
    for i in range(1, headers_range.Columns.Count + 1):

        values_in_column = [cell.Value or '' for cell in headers_range.Columns(i).Cells]
        values_list.append(' | '.join(values_in_column).replace('\n', ' '))

    WS_working.UsedRange.Rows(last_row_number).GetOffset(1, 0).Value = values_list # Fill with new header names
    headers_range.EntireRow.Delete() # remove previous header

    headers_range = WS_working.UsedRange.Rows(1)
    last_row_number = headers_range.Rows(headers_range.Rows.Count).Row



    if additional_column == True:

        new_column_range = WS_working.UsedRange.Columns(headers_range.Columns.Count).GetOffset(0, 1)

        headers_range = headers_range.GetResize(RowSize = headers_range.Rows.Count, ColumnSize = headers_range.Columns.Count + 1)

        new_column_range.value = os.path.basename(file_path).split('.')[0]
        new_column_range.Cells(1,1).value = ' | '.join(['Новая колонка' for i in range(number_of_level)])

        WS_working.UsedRange.Cells(last_row_number + 1,1).Copy()

        
        new_column_range.PasteSpecial(-4122)
        # new_column_range.NumberFormat = "General"
        ExcelApp.CutCopyMode = False


        headers_range.Cells(1,1).Copy()
        headers_range.PasteSpecial(-4122)
        ExcelApp.CutCopyMode = False

        new_column_range.Columns(1).Cut()

        WS_working.UsedRange.Columns(1).Insert(Shift= -4161)



    
    if original_sheet == False:
        WS_original.Delete()
        
    if informational_sheet == False:
        WS_informational.Delete()

    inner_style_headers(headers_range)
    

    file_name = os.path.splitext(os.path.basename(file_path))[0]
    return [WB_working], [file_name]
