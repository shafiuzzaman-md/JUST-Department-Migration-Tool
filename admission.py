import datetime
import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import pandas as pd
import pyodbc
import xlrd
from numpy.distutils.fcompiler import none
import logging

unit_name = none;
applicants_roll = 0;

root = tk.Tk()
result = ""
canvas1 = tk.Canvas(root, width=300, height=300, bg='lightsteelblue')
canvas1.pack()
connection_string = pyodbc.connect(
    'Driver={SQL Server};'
    'Server=.\sqlexpress;'
    'Database=Admission2019;'
    'Trusted_Connection=yes;')
cursor: object = connection_string.cursor()


def get_allotted_subject_id(position):
    cursor.execute(
        "SELECT AllottedDepartmentId FROM PassedApplicants WHERE Position ='%s'" % position)
    for row in cursor.fetchall():
        return row[0]


def get_department_status_by_id(department_id):
    cursor.execute(
        "SELECT  AllottedSeats FROM Departments WHERE Id =?", department_id)
    for row in cursor.fetchall():
        return row[0]


def get_non_confirmed_applicants(unit):
    logging.info(
        "Getting Applicants who did not confirmed admission of Unit " + unit)
    cursor.execute(
        f"SELECT * FROM [PassedApplicants] WHERE IsAdmissionConfirmed != 1 and [AllottedDepartmentOrder] > 0 and [AllottedDepartmentId] > 0 and UnitName ='{unit}'")
    columns = [column[0] for column in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def import_file():
    import_file_path = filedialog.askopenfilename()
    if import_file_path is '':
        return False, False
    df = pd.read_excel(import_file_path)
    wb = xlrd.open_workbook(import_file_path)
    ws = wb.sheet_by_index(0)
    position_list = ws.col_values(0)
    logging.info("position_list: " + str(position_list))
    roll_list = ws.col_values(1)
    logging.info("roll_list: " + str(roll_list))
    logging.info("length " + str(len(position_list)))
    return position_list, roll_list


def getExcelOfAdmissionConfirm():
    global applicants_roll
    position_list, roll_list = import_file()
    if not position_list:
        return False
    index = 1
    while index < len(position_list):
        logging.info("index " + str(index))
        pos = int(position_list[index])
        roll = int(roll_list[index])
        logging.info("Position: " + str(pos))
        logging.info("Roll" + str(roll))
        applicants_roll = str(applicants_roll) + "," + str(roll)
        str_unit = "'" + unit_name + "'"
        sql = "UPDATE PassedApplicants SET [IsAdmissionConfirmed] = 1 WHERE [Position] = " + str(
            pos) + " AND [Roll] = " + str(roll) + " AND [UnitName] = " + str_unit
        cursor.execute(sql)
        logging.info("Admission Confirmed")
        index = index + 1
    cursor.commit()
    non_confirmed_applicants = get_non_confirmed_applicants(unit_name)
    no_of_non_confirmed_applicants = len(non_confirmed_applicants)
    logging.info("Total: " + str(no_of_non_confirmed_applicants))
    for applicant in non_confirmed_applicants:
        sql = "UPDATE PassedApplicants SET [IsAdmissionCancelled] = 1 WHERE [Position] = " + str(
            applicant.get("Position")) + " AND [roll] = " + str(applicant.get("Roll")) + " AND [UnitName] = " + str_unit
        cursor.execute(sql)
        allotted_department_id = get_allotted_subject_id(applicant.get("Position"))
        allotted_seats = get_department_status_by_id(allotted_department_id)
        allotted_seats = allotted_seats - 1
        seat_status = True
        x = datetime.now()
        cursor.execute(
            "UPDATE Departments SET [SeatStatus] = ?, [AllottedSeats] = ?, [UpdatedDate] = ? WHERE [Id] = ?",
            seat_status, allotted_seats, x.strftime("%d%b%I%M%p"), allotted_department_id)
    cursor.commit()
    root.destroy()
    return True


def getExcelOfStopAutoMigrations():
    global applicants_roll
    position_list, roll_list = import_file()
    if not position_list:
        return False
    index = 1
    while index < len(position_list):
        logging.info("index " + str(index))
        pos = int(position_list[index])
        roll = int(roll_list[index])
        logging.info("Position: " + str(pos))
        logging.info("Roll" + str(roll))
        applicants_roll = str(applicants_roll) + "," + str(roll)
        str_unit = "'" + unit_name + "'"
        sql = "UPDATE PassedApplicants SET [IsAutoMigrationOff] = 1 WHERE [Position] = " + str(
            pos) + " AND [roll] = " + str(roll) + " AND [UnitName] = " + str_unit
        cursor.execute(sql)
        index = index + 1
    cursor.commit()
    root.destroy()
    return True


def getExcelOfAdmissionCancel():
    global applicants_roll
    position_list, roll_list = import_file()
    if not position_list:
        return False
    index = 1
    while index < len(position_list):
        logging.info("index " + str(index))
        pos = int(position_list[index])
        roll = int(roll_list[index])
        logging.info("Position: " + str(pos))
        logging.info("Roll" + str(roll))
        applicants_roll = str(applicants_roll) + "," + str(roll)
        str_unit = "'" + unit_name + "'"
        allotted_department_id = get_allotted_subject_id(pos)
        sql = "UPDATE PassedApplicants SET [IsAdmissionCancelled] = 1 WHERE [Position] = " + str(
            pos) + " AND [roll] = " + str(roll) + " AND [UnitName] = " + str_unit
        cursor.execute(sql)
        allotted_seats = get_department_status_by_id(allotted_department_id)
        allotted_seats = allotted_seats - 1
        seat_status = True
        x = datetime.now()
        cursor.execute(
            "UPDATE Departments SET [SeatStatus] = ?, [AllottedSeats] = ?, [UpdatedDate] = ? WHERE [Id] = ?",
            seat_status, allotted_seats, x.strftime("%d%b%I%M%p"), allotted_department_id)
        index = index + 1
    cursor.commit()
    root.destroy()
    return True


def process_admission(unit, admission_command):
    drive = "F"
    os.chdir(drive + ':')  # change directory
    path: object = "Unit_" + unit + "_process_admission" + str(datetime.now().strftime("_%d_%b_%I_%M_%p"))
    if not os.path.exists(path):
        os.makedirs(path)
    os.chdir(path)
    logging.basicConfig(filename="info.log", format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S',
                        level=logging.INFO)
    global unit_name, applicants_roll
    unit_name = unit
    if admission_command == "confirm":
        browseButton_Excel = tk.Button(text='Import Excel File of Confirmed Applicants',
                                       command=getExcelOfAdmissionConfirm,
                                       bg='green',
                                       fg='white',
                                       font=('helvetica', 10, 'bold'))

    if admission_command == "stop_migration":
        browseButton_Excel = tk.Button(text='Import Excel File of Stop Auto Migrations',
                                       command=getExcelOfStopAutoMigrations, bg='blue',
                                       fg='white',
                                       font=('helvetica', 10, 'bold'))

    if admission_command == "cancel":
        browseButton_Excel = tk.Button(text='Import Excel File of Admission Cancel', command=getExcelOfAdmissionCancel,
                                       bg='red',
                                       fg='white',
                                       font=('helvetica', 10, 'bold'))

    canvas1.create_window(150, 150, window=browseButton_Excel)
    root.mainloop()
    return True, applicants_roll
