import pyodbc

connection_string = pyodbc.connect(
    'Driver={SQL Server};'
    'Server=.\sqlexpress;'
    'Database=Admission2019;'
    'Trusted_Connection=yes;')
cursor: object = connection_string.cursor()

# Create PassedApplicants table
create_passed_applicants_table = 'CREATE TABLE PassedApplicants([Id] [int] NOT NULL, [Position] [real] NULL, [ApplicationId] [nvarchar](200) NOT NULL, [Roll] [nvarchar](max) NULL,[UnitName] [nvarchar](200) NOT NULL,[TotalMarks] [real] NULL,[PhoneNumber] [nvarchar](max) NULL,[Name] [nvarchar](max) NULL,[Father] [nvarchar](max) NULL,[Photo] [nvarchar](max) NULL,[QuotaName] [nvarchar](450) NULL, [AllottedDepartment] [nvarchar](450) NULL, [UpdatedDate] [nvarchar](450) NULL, [IsAutoMigrationOff] [bit] null, [IsAdmissionCancelled] [bit] null, [Remarks] [nvarchar](450) NULL,)'
cursor.execute(create_passed_applicants_table)

# Create Departments table
create_departments_table = 'CREATE TABLE Departments([Id] [int] NOT NULL,[Code] [nvarchar](200) NOT NULL,[DepartmentName] [nvarchar](200) NOT NULL,[UnitName] [nvarchar](10) NULL,[TotalSeats] [int], [AllottedSeats] [int], [SeatStatus] [bit], [StartingPosition] [real] NULL, [EndingPosition] [real] NULL,)'
cursor.execute(create_departments_table)

connection_string.commit()

# Copy data from Applications table
copy_from_applications_table = "SELECT Id, UserId, ApplicationId, [Position], Roll, UnitName, TotalMarks FROM Admission2019.Applications WHERE Position IS NOT NULL"
cursor.execute(copy_from_applications_table)
columns = [column[0] for column in cursor.description]
applications = [dict(zip(columns, row)) for row in cursor.fetchall()]

# Copy data from Users table
copy_from_users_table = "SELECT Id, PhoneNumber, [Name], Father, Photo, QuotaName FROM Admission2019.AspNetUsers"
cursor.execute(copy_from_users_table)
columns = [column[0] for column in cursor.description]
users = [dict(zip(columns, row)) for row in cursor.fetchall()]

# Copy data from Subjects table
copy_from_subjects_table = "SELECT * FROM Admission2019.Subjects"
cursor.execute(copy_from_subjects_table)
columns = [column[0] for column in cursor.description]
subjects = [dict(zip(columns, row)) for row in cursor.fetchall()]

# insert data into PassedApplicants table
for application in applications:
    for user in users:
        if application.get("UserId") == user.get("Id"):
            cursor.execute(
                "INSERT INTO PassedApplicants([Id],[Position],[ApplicationId],[Roll],[UnitName],[TotalMarks],[PhoneNumber],[Name],[Father],[Photo],[QuotaName],[IsAutoMigrationOff], [IsAdmissionCancelled]) values (?,?, ?, ?,?, ?, ?,?, ?, ?,?, ?,?)",
                application.get('Id'), application.get('Position'), application.get('ApplicationId'), application.get('Roll'),
                application.get('UnitName'), application.get('TotalMarks'), user.get('PhoneNumber'), user.get('Name'),
                user.get('Father'), user.get('Photo'), user.get('QuotaName'), 0, 0)
connection_string.commit()

# insert data into Departments table
for subject in subjects:
    cursor.execute(
        "INSERT INTO Departments ([Id],[Code],[DepartmentName],[UnitName],[TotalSeats],[AllottedSeats],[SeatStatus]) values (?,?,?,?,?,?,?)",
        subject.get("Id"), subject.get("Code"), subject.get("Name"), subject.get("UniteName"), 0, 0, 1)
connection_string.commit()
