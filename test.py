import pandas as pd
from sqlalchemy import create_engine
import os
from pathlib import Path
import shutil
from Procedure_Excel import ProcExec

# File path
file_path = r' ' # r'D:\Заливка Плана РМ\Ресурсна Модель ( РЦ Всі склади ) Прогноз (лютий, 2023)1.xlsx'
directory = Path("//fozzy.lan/Documents/Holding/Departments/Silpo/Логистика_GERMES/Exchange/Staf для заливки/План на месяц/")
file_name = "Ресурсна Модель ( РЦ Всі склади ) Прогноз"
file_extension = ".xlsx"
file_name_found = None
# Задайте путь к каталогу для копирование файла
destination_folder = r'D:/Заливка Плана РМ/'

for file in directory.glob(f"{file_name}*{file_extension}"):
    file_path = file
    file_name_found = file.name

if os.access(file_path, os.F_OK) == True:
    # Переместите файл в новое место
    new_location = shutil.copy2(file_path, destination_folder + file_name_found)

    # Распечатать новое расположение файла
    print("{0} перемещается в нужное место, {1}".format(file_path, new_location))

    sheets = {
        "Kvitneve |Result_Forecast_Data": 2082,
        "Peremoga | Result_Forecast_Data": 2081,
        "Lviv | Result_Forecast_Data": 2930,
        "Odesa |Result_Forecast_ Data": 2102,
        "Zaporizhja|Result_Forecast_Data": 2135
    }
    # Set up SQL connection
    engine = create_engine('mssql+pyodbc://s-kv-center-s64/PlanRC?driver=SQL+Server')

    # Цикл для обработки данных по каждому листу
    for sheet, idrc in sheets.items():

        # Read the excel file and set the correct sheet name and header rows
        df = pd.read_excel(file_path, sheet_name=sheet, header=4)
        # Извлечение названий столбцов из датафрейма
        column_names = df.columns.tolist()
        # Создание нового датафрейма с названиями столбцов
        columns_df = pd.DataFrame(column_names, columns=["column_name"])

        # добавление столбца счетчика строк до загрузки в SQL
        columns_df = columns_df.iloc[1:, ]
        columns_df['ID'] = [i for i in range(1, len(columns_df) + 1)]
        columns_df['RCID'] = idrc

        # Создание нового датафрейма с названиями столбцов
        df_type= pd.read_excel(file_path, sheet_name=sheet, header=2)
        # Извлечение названий столбцов из датафрейма
        column_type = df_type.columns.tolist()
        # добавление столбца  строк до загрузки в SQL
        typpe = pd.DataFrame(column_type, columns=["Type"])
        column_type = df_type.iloc[1:, ]
        columns_df['Type'] = typpe

        # Создание нового датафрейма с названиями столбцов
        df_Total = pd.read_excel(file_path, sheet_name=sheet, header=0)
        column_Total = df_Total.columns.tolist()
        Total_Data = pd.DataFrame(column_Total, columns=["Total_Data"])
        columns_df['Total_Data'] = Total_Data

        # Загрузка названий столбцов в SQL
        columns_df.to_sql("_Ant_PlanRM_Column_Name", engine, if_exists="replace", index=False)

        # Загрузка данных в SQL таблицу
        df.to_sql("_Ant_PlanRM_Data_ImportExcel", engine, if_exists="replace", index=False)

        p = ProcExec(r'D:\Заливка Плана РМ\проверка_данных.xlsm')
        p.run_macro("Module1.Execute_command")
        p.save_and_close()


        print(sheet)

else:
    print("Файл не найден")