import tkinter as tk
from tkinter import filedialog
import pandas as pd
import pyodbc
import xlrd

root = tk.Tk()

canvas1 = tk.Canvas(root, width=300, height=300, bg='lightsteelblue')
canvas1.pack()
connection_string = pyodbc.connect(
    'Driver={SQL Server};'
    'Server=.\sqlexpress;'
    'Database=Admission2019;'
    'Trusted_Connection=yes;')
cursor: object = connection_string.cursor()


def getExcel():
    global df

    import_file_path = filedialog.askopenfilename()
    df = pd.read_excel(import_file_path)
    wb = xlrd.open_workbook(import_file_path)
    ws = wb.sheet_by_index(0)
    position_list = ws.col_values(0)
    roll_list = ws.col_values(1)
    index = 1
    print("length " + str (len(position_list)))
    while index < len(position_list):
        print("index " + str(index))
        pos = int(position_list[index])
        roll = int(roll_list[index])
        print(pos)
        print(roll)
        sql = "UPDATE PassedApplicants SET [IsAdmissionConfirmed] = 1 WHERE [Position] = " + str(pos) + " AND [roll] = " + str(roll)
        cursor.execute(sql)
        index = index + 1
    cursor.commit()


browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=browseButton_Excel)

root.mainloop()
