import datetime
import pyodbc
from openpyxl import Workbook

applicants = None
subject_choices = None
connection_string = pyodbc.connect(
    'Driver={SQL Server};'
    'Server=.\sqlexpress;'
    'Database=Admission2019;'
    'Trusted_Connection=yes;')
cursor = connection_string.cursor()


def backup_db() -> object:
    print("Database backup started...")
    connection_string.autocommit = True
    backup_file = "'d:\\Admission2019" + str(datetime.datetime.now().strftime("_%d_%B_%I_%M_%p")) + ".bak'"
    sql = "BACKUP DATABASE [Admission2019] TO DISK = N" + backup_file
    cursor.execute(sql)
    connection_string.autocommit = False
    print("Find the backup file in " + backup_file)
    print("Database backup finished...")


def get_applicants(unit):
    print("Getting Applicants who did not cancelled admission and did not initiate Auto Migration OFF of Unit " + unit)
    cursor.execute(
        "SELECT * FROM PassedApplicants WHERE [IsAdmissionCancelled] != 1 and [IsAutoMigrationOff] != 1 and UnitName ='%s'" % unit)
    columns = [column[0] for column in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def get_first_position(unit):
    print("Getting 1st Applicant...")
    cursor.execute(
        "SELECT MIN(Position) FROM PassedApplicants WHERE [IsAdmissionCancelled] != 1 and [IsAutoMigrationOff] != 1 and UnitName ='%s'" % unit)
    for row in cursor.fetchall():
        return row[0]


def get_subject_choices_by_id(application_id):
    cursor.execute(
        "SELECT SubjectId, [Order], ApplicationId FROM Admission2019.SubjectChoices WHERE ApplicationId ='%s'" % application_id)
    columns = [column[0] for column in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def get_subject_id_by_order(application_table_id, order):
    cursor.execute("SELECT SubjectId FROM Admission2019.SubjectChoices WHERE ApplicationId = ? AND [Order] = ?",
                   application_table_id, order)
    for row in cursor.fetchall():
        return row[0]


def get_department_status_by_id(department_id):
    cursor.execute(
        "SELECT SeatStatus, TotalSeats, AllottedSeats, DepartmentName FROM Departments WHERE Id =?", department_id)
    for row in cursor.fetchall():
        return row[0], row[1], row[2], row[3]


def get_departments():
    cursor.execute("SELECT * FROM Departments WHERE SeatStatus = 1")
    columns = [column[0] for column in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def allocate_subject(application_table_id, applicant_position):
    choices = get_subject_choices_by_id(application_table_id)
    number_of_choices = len(choices)
    print("No. of Choices: " + str(number_of_choices))
    x = datetime.datetime.now()
    # Make IsAdmissionCancelled True for the applicant who did not fill up the choice form
    if number_of_choices == 0:
        cursor.execute("UPDATE [PassedApplicants] SET [IsAdmissionCancelled] = ?, [UpdatedDate] = ? WHERE [Id] = ?",
                       1, x.strftime("%d%b%I%M%p"), application_table_id)
        print("Admission is being cancelled as ")
        return "did not fill up the choice form"
    order = 1
    while order <= number_of_choices:
        subject_id = get_subject_id_by_order(application_table_id, order)
        seat_status, total_seats, allotted_seats, department_name = get_department_status_by_id(subject_id)
        print(
            str(order) + ": " + str(department_name) + " Total Seats: " + str(total_seats) + " Allotted Seats: " + str(
                allotted_seats))
        if seat_status is True and allotted_seats <= total_seats and total_seats is not 0:
            cursor.execute("UPDATE [PassedApplicants] SET [AllottedDepartment] = ?, [UpdatedDate] = ? WHERE [Id] = ?",
                           department_name, x.strftime("%d%b%I%M%p"), application_table_id)
            allotted_seats = allotted_seats + 1
            if allotted_seats == 0:
                cursor.execute("UPDATE Departments SET [AllottedSeats] = ?, [StartingPosition] = ? WHERE [Id] = ?",
                               allotted_seats, applicant_position, subject_id)
            elif allotted_seats == total_seats:
                cursor.execute(
                    "UPDATE Departments SET [AllottedSeats] = ?, [EndingPosition] = ?, [SeatStatus] = ? WHERE [Id] = ?",
                    allotted_seats, applicant_position, 0, subject_id)
            else:
                cursor.execute("UPDATE Departments SET [AllottedSeats] = ? WHERE [Id] = ?",
                               allotted_seats, subject_id)
            return department_name
        else:
            order = order + 1
    return "No Department"


def get_applicant_id_by_position(applicant_position):
    cursor.execute("SELECT Id FROM PassedApplicants WHERE Position = '%d'" % applicant_position)
    for row in cursor.fetchall():
        return row[0]


def read_data_from_db(query):
    print("Reading data from db......")
    applicants_data = cursor.execute(query)
    return applicants_data


def write_applicants_data_to_excel(applicants_data):
    print("Writing to excel......")
    wb = Workbook(write_only=True)
    migration_ws = wb.create_sheet("Migration")
    # write header
    migration_ws.append(["Position", "Name", "ApplicationId", "Roll", "Department", "Unit", "Phone", "Quota"
                         "IsAutoMigrationOff"])

    # write data
    for applicant in applicants_data:
        position = applicant[1]
        name = applicant[7]
        phone = applicant[6]
        applicant_id = applicant[2]
        roll = applicant[3]
        department = applicant[11]
        unit = applicant[4]
        quota = applicant[10]
        auto_migration_off = applicant[13]
        migration_ws.append([position, name, applicant_id, roll, department, unit, phone, quota, auto_migration_off])
    migration_result = "d:\\Migration" + str(datetime.datetime.now().strftime("_%d_%B_%I_%M_%p")) + ".xlsx"
    wb.save(migration_result)
    print("Find excel file into " + migration_result)


unit_name = input("Unit(Press A/B/C/D/E/F): ")
unit_name = unit_name.upper()
if unit_name == 'A' or unit_name == 'B' or unit_name == 'C' or unit_name == 'D' or unit_name == 'E' or unit_name == 'F':
    while True:
        try:
            end_position = int(input('End Position: '))
            break
        except:
            print("Invalid Input!")
    backup_db()
    applicants = get_applicants(unit_name)
    position = get_first_position(unit_name)
    print("Starting Position: " + str(position))
    no_of_applicants = len(applicants)
    print("Total: " + str(no_of_applicants))
    while position <= end_position:
        print("---------------------------------" + str(no_of_applicants) + " remains ------------------------------")
        print("Migration Started for Position " + str(position))
        applicant_id = get_applicant_id_by_position(position)
        print("Id: " + str(applicant_id))
        if applicant_id is not None:
            department = allocate_subject(applicant_id, position)
            print("Position " + str(position) + " " + department)
            cursor.commit()  # Commit db changes
        position = position + 1
        no_of_applicants = no_of_applicants - 1
        if no_of_applicants == 0:
            break

    print("Exporting migration result into excel....")
    get_applicants_query = " SELECT * FROM PassedApplicants WHERE IsAdmissionCancelled = 0 "
    migration_data = read_data_from_db(get_applicants_query)
    write_applicants_data_to_excel(applicants_data=migration_data)

    print("Exporting department status into excel....")
    get_departments_query = " SELECT * FROM Departments"
    department_data = read_data_from_db(get_departments_query)
    write_data_to_excel(department_data=migration_data)
else:
    print("Invalid Unit Name")
