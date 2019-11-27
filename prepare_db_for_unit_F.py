def prepare_database_F(connection_string, unit, group):
    cursor = connection_string.cursor()
    # Create PassedApplicants table
    create_passed_applicants_table = 'CREATE TABLE [PassedApplicants' + unit + group + ']([Id] [int] NOT NULL, [Position] [real] NULL, [ApplicationId] [nvarchar](200) NOT NULL, [Roll] [nvarchar](max) NULL,[UnitName] [nvarchar](200) NOT NULL,[TotalMarks] [real] NULL,[PhoneNumber] [nvarchar](max) NULL,[Name] [nvarchar](max) NULL,[Father] [nvarchar](max) NULL,[Photo] [nvarchar](max) NULL,[QuotaName] [nvarchar](450) NULL,[AllottedDepartmentId] [int] NULL, [AllottedDepartment] [nvarchar](450) NULL, [AllottedDepartmentOrder] [int] NULL, [UpdatedDate] [nvarchar](450) NULL, [IsAdmissionConfirmed] [bit] null, [IsAutoMigrationOff] [bit] null, [IsAdmissionCancelled] [bit] null, [Remarks] [nvarchar](450) NULL,)'
    try:
        cursor.execute(create_passed_applicants_table)
    except:
        result = "Applicants DB is already updated!"
        return result

    # Create Departments table
    create_departments_table = 'CREATE TABLE [Departments' + unit + group + ']([Id] [int] NOT NULL,[Code] [nvarchar](200) NOT NULL,[DepartmentName] [nvarchar](200) NOT NULL,[UnitName] [nvarchar](10) NULL,[TotalSeats] [int], [AllottedSeats] [int], [SeatStatus] [bit], [UpdatedDate] [nvarchar](450) NULL)'
    cursor.execute(create_departments_table)

    connection_string.commit()
    unit_name = "'" + unit + "'"
    # Copy data of applicants who have filled up choice form
    # copy_from_applications_table = "SELECT Id, UserId, ApplicationId, [Position], Roll, UnitName, TotalMarks FROM Admission2019.Applications WHERE UnitName = " + unit_name + " AND Position IS NOT NULL AND Admission2019.Applications.Id in (SELECT ApplicationId From Admission2019.SubjectChoices Where [Order] > 0) ORDER BY UnitName asc, Position asc"
    if group == 'BUSINESS STUDIES':
        copy_from_applications_table = "SELECT Admission2019.Applications.Id, UserId, ApplicationId, [Position], Roll, UnitName, TotalMarks  FROM  Admission2019.Applications INNER JOIN Admission2019.AspNetUsers ON Admission2019.Applications.UserId = Admission2019.AspNetUsers.Id WHERE UnitName = " + unit_name + " AND  Admission2019.AspNetUsers.HSCGroup != 'SCIENCE' AND Admission2019.AspNetUsers.HSCGroup != 'HUMANITIES' AND Position IS NOT NULL AND Admission2019.Applications.Id in (SELECT ApplicationId From Admission2019.SubjectChoices Where[Order] > 0) ORDER BY UnitName asc, Position asc"
    else:
        copy_from_applications_table = "SELECT Admission2019.Applications.Id, UserId, ApplicationId, [Position], Roll, UnitName, TotalMarks  FROM  Admission2019.Applications INNER JOIN Admission2019.AspNetUsers ON Admission2019.Applications.UserId = Admission2019.AspNetUsers.Id WHERE UnitName = " + unit_name + " AND  Admission2019.AspNetUsers.HSCGroup = " + "'" + group + "'" "AND Position IS NOT NULL AND Admission2019.Applications.Id in (SELECT ApplicationId From Admission2019.SubjectChoices Where[Order] > 0) ORDER BY UnitName asc, Position asc"
    cursor.execute(copy_from_applications_table)
    columns = [column[0] for column in cursor.description]
    applications = [dict(zip(columns, row)) for row in cursor.fetchall()]

    # Copy data from Users table
    copy_from_users_table = "SELECT Id, PhoneNumber, [Name], Father, Photo, QuotaName FROM Admission2019.AspNetUsers"
    cursor.execute(copy_from_users_table)
    columns = [column[0] for column in cursor.description]
    users = [dict(zip(columns, row)) for row in cursor.fetchall()]

    # Copy data from Subjects table
    copy_from_subjects_table = "SELECT * FROM Admission2019.Subjects WHERE  UniteName = " + unit_name
    cursor.execute(copy_from_subjects_table)
    columns = [column[0] for column in cursor.description]
    subjects = [dict(zip(columns, row)) for row in cursor.fetchall()]

    # insert data into PassedApplicants table
    for application in applications:
        for user in users:
            if application.get("UserId") == user.get("Id"):
                cursor.execute(
                    "INSERT INTO [PassedApplicants" + unit + group + "] ([Id],[Position],[ApplicationId],[Roll],[UnitName],[TotalMarks],[PhoneNumber],[Name],[Father],[Photo],[QuotaName], [AllottedDepartmentId], [AllottedDepartmentOrder], [IsAdmissionConfirmed], [IsAutoMigrationOff], [IsAdmissionCancelled]) values (?,?, ?, ?,?, ?, ?,?, ?, ?,?, ?,?,?,?,?)",
                    application.get('Id'), application.get('Position'), application.get('ApplicationId'),
                    application.get('Roll'),
                    application.get('UnitName'), application.get('TotalMarks'), user.get('PhoneNumber'),
                    user.get('Name'),
                    user.get('Father'), user.get('Photo'), user.get('QuotaName'), 0, 0, 0, 0, 0)
    connection_string.commit()

    # insert data into Departments table
    for subject in subjects:
        cursor.execute(
            "INSERT INTO [Departments" + unit + group + "] ([Id],[Code],[DepartmentName],[UnitName],[TotalSeats],[AllottedSeats],[SeatStatus]) values (?,?,?,?,?,?,?)",
            subject.get("Id"), subject.get("Code"), subject.get("Name"), subject.get("UniteName"), 0, 0, 1)
    connection_string.commit()
    result = "Applicants DB is updated successfully!"
    return result
