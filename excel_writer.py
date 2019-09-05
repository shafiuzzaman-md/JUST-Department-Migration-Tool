
from openpyxl import Workbook

import datetime

import pandas as pd
import pyodbc

connection_string = pyodbc.connect(
    'Driver={SQL Server};'
    'Server=.\sqlexpress;'
    'Database=Admission2019;'
    'Trusted_Connection=yes;')
cursor = connection_string.cursor()


def read_data_from_db():
    query = " select * from PassedApplicants "
    applicants_data = cursor.execute(query)
    return applicants_data


def write_to_default_ws_row_col(applicants):
    wb = Workbook(write_only=True)
    migration_ws = wb.create_sheet("Migration")
    # write header
    migration_ws.append(["Id", "AllottedDepartment"])

    # write data
    for applicant in applicants:
        applicant_id = applicant[0]
        applicant_department = applicant[1]
        migration_ws.append([applicant_id, applicant_department])
    migration_result = "d:\\Migration" + str(datetime.datetime.now().strftime("_%d_%B_%I_%M_%p")) + ".xlsx"
    wb.save(migration_result)


if __name__ == "__main__":
    migration_data = read_data_from_db()
    write_to_default_ws_row_col(applicants=migration_data)
