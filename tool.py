import pyodbc
import wx

from admission import process_admission
from migration import execute_migration
from prepare_database_for_migration import prepare_database
from prepare_db_for_unit_F import prepare_database_F
from migration_F import execute_migration_F

APP_EXIT = 1
unit = ""
group = ""
driver = 'Driver={SQL Server};'
sever = 'Server=.\sqlexpress;'
db_name = 'Database=DB_A4A307_Migration_Test2;'
auth = 'Trusted_Connection=yes;'
con = driver + sever + db_name + auth
connection_string = pyodbc.connect(con)


class UI(wx.Frame):

    def __init__(self, *args, **kwargs):
        super(UI, self).__init__(*args, **kwargs)
        pnl = wx.Panel(self)
        self.sb = self.CreateStatusBar(1)
        # l1 = wx.StaticText(pnl, -1, "Connection String:", pos=(10, 10), size=(200, 50))
        # l3 = wx.StaticText(pnl, -1, str(con), pos=(130, 10), size=(800, 150))

        self.rb1 = wx.RadioButton(pnl, label='Unit A', pos=(150, 90))
        self.rb2 = wx.RadioButton(pnl, label='Unit B', pos=(300, 90))
        self.rb3 = wx.RadioButton(pnl, label='Unit C', pos=(450, 90))
        self.rb4 = wx.RadioButton(pnl, label='Unit D', pos=(150, 120))
        self.rb5 = wx.RadioButton(pnl, label='Unit E', pos=(300, 120))
        self.rb6 = wx.RadioButton(pnl, label='Unit F (BUS)', pos=(450, 120))
        self.rb7 = wx.RadioButton(pnl, label='Unit F (SCI)', pos=(600, 90))
        self.rb8 = wx.RadioButton(pnl, label='Unit F (HUM)', pos=(600, 120))

        self.prepareDatabase = wx.Button(pnl, label='Get Passed Applicants for Subject Allocation', pos=(230, 200),
                                         size=(300, 35))
        self.prepareDatabase.SetBackgroundColour(wx.Colour(255, 255, 255))

        self.confirmAdmissionButton = wx.Button(pnl, label='Confirm Admissions', pos=(230, 250), size=(300, 35))
        self.confirmAdmissionButton.SetBackgroundColour(wx.Colour(255, 255, 255))

        self.stopAutoMigrationButton = wx.Button(pnl, label='Stop Auto Migrations', pos=(230, 300), size=(300, 35))
        self.stopAutoMigrationButton.SetBackgroundColour(wx.Colour(255, 255, 255))

        self.cancelAdmissionButton = wx.Button(pnl, label='Cancel Admission', pos=(230, 350), size=(300, 35))
        self.cancelAdmissionButton.SetBackgroundColour(wx.Colour(255, 255, 255))

        self.migrationButton = wx.Button(pnl, label='Subject Allocation and Migration', pos=(230, 400), size=(300, 35))
        self.migrationButton.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.InitUI()

    def InitUI(self):

        self.rb1.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb2.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb3.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb4.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb5.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb6.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb7.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb8.Bind(wx.EVT_RADIOBUTTON, self.SetVal)

        self.prepareDatabase.Bind(wx.EVT_BUTTON, self.OnprepareDatabaseButtonPressed)
        self.confirmAdmissionButton.Bind(wx.EVT_BUTTON, self.OnConfirmAdmissionButtonPressed)
        self.stopAutoMigrationButton.Bind(wx.EVT_BUTTON, self.OnStopAutoMigrationButtonPressed)
        self.cancelAdmissionButton.Bind(wx.EVT_BUTTON, self.OnCancelAdmissionButton)
        self.migrationButton.Bind(wx.EVT_BUTTON, self.OnMigrationButtonPressed)

        self.SetSize((800, 600))
        self.Centre()

    def OnprepareDatabaseButtonPressed(self, e):
        if unit == 'F':
            result = prepare_database_F(connection_string, unit, group)
        else:
            result = prepare_database(connection_string, unit)
        wx.MessageBox(result, 'Info', wx.OK | wx.ICON_INFORMATION)

    def OnConfirmAdmissionButtonPressed(self, e):
        result, positions = process_admission(unit, "confirm")
        msg = unit + " unit admissions are confirmed successfully for merit position" + str(positions)
        wx.MessageBox(msg, 'Info', wx.OK | wx.ICON_INFORMATION)
        self.Close()

    def OnStopAutoMigrationButtonPressed(self, e):
        result, positions = process_admission(unit, "stop_migration")
        if result is not None:
            msg = unit + " unit stop auto migration requests are executed successfully for position " + str(
                positions)
            wx.MessageBox(msg, 'Info', wx.OK | wx.ICON_INFORMATION)
            self.Close()

    def OnCancelAdmissionButton(self, e):
        result, positions = process_admission(unit, "cancel")
        msg = unit + " unit admissions are cancelled successfully of position " + str(positions)
        wx.MessageBox(msg, 'Info', wx.OK | wx.ICON_INFORMATION)
        self.Close()

    def OnMigrationButtonPressed(self, e):
        if unit == 'F':
            result = execute_migration_F(unit, group)
        else:
            result = execute_migration(unit)
        if result is not None:
            wx.MessageBox(result, 'Info', wx.OK | wx.ICON_INFORMATION)
            self.Close()

    def OnQuit(self, e):
        self.Close()

    def SetVal(self, e):
        global unit
        global group
        if self.rb1.GetValue():
            unit = 'A'
        if self.rb2.GetValue():
            unit = 'B'
        if self.rb3.GetValue():
            unit = 'C'
        if self.rb4.GetValue():
            unit = 'D'
        if self.rb5.GetValue():
            unit = 'E'
        if self.rb6.GetValue():
            unit = 'F'
            group = 'BUSINESS STUDIES'
        if self.rb7.GetValue():
            unit = 'F'
            group = 'SCIENCE'
        if self.rb8.GetValue():
            unit = 'F'
            group = 'HUMANITIES'
        msg = "Unit " + unit + " is Selected"
        self.sb.SetStatusText(msg, 0)


def main( ):
    app = wx.App()
    ex = UI(None, title='JUST Migration Tool')
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
