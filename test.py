import pyodbc

# Connection String
connection_string = pyodbc.connect(
    'Driver={SQL Server};'
    'Server=.\sqlexpress;'
    'Database=DB_A4A307_Production_Migration_test;'
    'Trusted_Connection=yes;')

cursor = connection_string.cursor()


def get_applicants(unit):
    cursor.execute(
        'SELECT * FROM PassedApplicants WHERE [IsAdmissionCancelled] != 1 and [IsAutoMigrationOff] != 1 and UnitName =\'%s\'' % unit)
    columns = [column[0] for column in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


applicants = get_applicants('A')
print(applicants)