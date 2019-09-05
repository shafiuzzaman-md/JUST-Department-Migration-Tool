import datetime
import sys
import pyodbc
import pandas as pd
from click._compat import raw_input

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=.\sqlexpress;'
                      'Database=Admission2019;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()
applications = None
subjectChoices = None
subjects = None
unit_name = 'A'
end_position = 10
table_name = "Migration" + str(datetime.datetime.now().strftime("%m_%d_%Y_%H_%M_%S")) + "Unit_" + unit_name


def get_applications():
    global applications
    cursor.execute(
        "SELECT Id, ApplicationId, Position, Roll, UnitName, AutoMigrationOFF FROM Admission2019.Applications")
    columns = [column[0] for column in cursor.description]
    applications = [dict(zip(columns, row)) for row in cursor.fetchall()]


def get_subject_choices():
    global subjectChoices
    cursor.execute('SELECT SubjectId, [Order], ApplicationId FROM Admission2019.SubjectChoices')
    columns = [column[0] for column in cursor.description]
    subjectChoices = [dict(zip(columns, row)) for row in cursor.fetchall()]


def get_subjects():
    global subjects
    cursor.execute('SELECT * FROM Admission2019.Subjects')
    columns = [column[0] for column in cursor.description]
    subjects = [dict(zip(columns, row)) for row in cursor.fetchall()]


def create_migration_table():
    create_table = "CREATE TABLE " + table_name + "([Id] [int] NULL, [Position] [int] NULL, [ApplicationId] [nvarchar](200) NULL, [RollNumber] [nvarchar](10) NULL, [Subject] [nvarchar](200) NULL, [Unit] [nvarchar](10) NULL)"
    cursor.execute(create_table)
    cursor.commit()


def write_to_excel():
    script = "SELECT * FROM " + table_name;
    cursor.execute(script)
    df = pd.read_sql_query(script, conn)
    name = 'D:\\' + table_name + '.xlsx'
    writer = pd.ExcelWriter(name)
    df.to_excel(writer, sheet_name='bar')
    writer.save()

def get_vacant_seats_by_subject_id(subject_id):
    global subjects
    for subject in subjects:
        if subject.get('Id') == subject_id:
            total = subject.get('TotalSeats')
            allotted = subject.get('AllottedSeats')
            return total, allotted


def set_subject(order, application_table_id):
    global subjectChoices
    for choice in subjectChoices:
        if choice.get('ApplicationId') == application_table_id:
            if choice.get('Order') == order:
                subject_id = choice.get('SubjectId')
                total_seats, allotted_seats = get_vacant_seats_by_subject_id(subject_id)
                if allotted_seats is None:
                    allotted_seats = 0
                vacant_seat = total_seats - allotted_seats
                if vacant_seat > 0:
                    allotted_seats += 1
                    cursor.execute("UPDATE [Admission2019].[Subjects] SET [AllottedSeats] = ? WHERE [Id] = ?",
                                   allotted_seats, subject_id)
                    return subject_id
                else:
                    order += 1
                    continue
    subject_id = None
    return subject_id


def allot_subject(application_table_id):
    global subjectChoices
    for choice in subjectChoices:
        if choice.get('ApplicationId') == application_table_id:
            return set_subject(1, application_table_id)


def get_subject_name_by_id(subject_id):
    global subjects
    for subject in subjects:
        if subject.get('Id') == subject_id:
            return subject.get('Name')


def get_application_table_id_by_position(position):
    for application in applications:
        if application.get("UnitName") == unit_name and application.get('Position') == position:
            #if application.get("Id") is None:
             #   position += 1
             #   get_application_table_id_by_position(position)
            return application.get("Id")


def start_migration(position):
    get_subjects()
    if position == end_position:
        return True
    for application in applications:
        application_table_id = get_application_table_id_by_position(position)
        application_id = application.get("ApplicationId")
        roll = application.get("Roll")
        auto_migration_off = application.get("AutoMigrationOFF")
        if auto_migration_off == 1:
            position += 1
        allotted_subject_id = allot_subject(application_table_id)
        if allotted_subject_id is not None:
            cursor.execute("UPDATE [Admission2019].[Applications] SET [SubjectId] = ? WHERE [Id] = ?",
                           allotted_subject_id, application_table_id)
        subject_name = get_subject_name_by_id(allotted_subject_id)
        if roll is None:
            roll = 0
        if subject_name is None:
            subject_name = "No Subject"
        if application_table_id is None:
            application_table_id = 0
        insert_into_table = "INSERT INTO dbo." + table_name + " (Id, Position, ApplicationId, RollNumber, [Subject], Unit) values (" + str(
            application_table_id) + "," + str(position) + ",'" + str(application_id) + "'," + str(
            roll) + ",'" + subject_name + "','" + str(unit_name) + "')"
        cursor.execute(insert_into_table)
        cursor.commit()
        # get_applications()
        position += 1
        start_migration(position)
    position += 1
    start_migration(position)


def backup_db():
    conn.autocommit = True
    backup_file = "BACKUP DATABASE [Admission2019] TO DISK = N'd:\\Admission2019" + str(
        datetime.datetime.now().strftime("%m_%d_%Y_%H_%M_%S")) + ".bak'"
    conn.cursor().execute(backup_file)
    conn.autocommit = False


backup_db()
create_migration_table()
get_applications()
get_subjects()
get_subject_choices()
print(start_migration(1))

