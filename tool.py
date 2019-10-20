import wx

from admission import process_admission
from migration import execute_migration

APP_EXIT = 1
unit = ""


class UI(wx.Frame):

    def __init__(self, *args, **kwargs):
        super(UI, self).__init__(*args, **kwargs)
        pnl = wx.Panel(self)
        self.sb = self.CreateStatusBar(1)

        self.rb1 = wx.RadioButton(pnl, label='Unit A', pos=(100, 10))
        self.rb2 = wx.RadioButton(pnl, label='Unit B', pos=(300, 10))
        self.rb3 = wx.RadioButton(pnl, label='Unit C', pos=(500, 10))
        self.rb4 = wx.RadioButton(pnl, label='Unit D', pos=(100, 40))
        self.rb5 = wx.RadioButton(pnl, label='Unit E', pos=(300, 40))
        self.rb6 = wx.RadioButton(pnl, label='Unit F', pos=(500, 40))
        self.confirmAdmissionButton = wx.Button(pnl, label='Confirm Admissions', pos=(250, 100), size=(200, 35))
        self.confirmAdmissionButton.SetBackgroundColour(wx.Colour(255, 255, 255))

        self.stopAutoMigrationButton = wx.Button(pnl, label='Stop Auto Migrations', pos=(250, 170), size=(200, 35))
        self.stopAutoMigrationButton.SetBackgroundColour(wx.Colour(255, 255, 255))

        self.cancelAdmissionButton = wx.Button(pnl, label='Cancel Admission', pos=(250, 240), size=(200, 35))
        self.cancelAdmissionButton.SetBackgroundColour(wx.Colour(255, 255, 255))

        self.migrationButton = wx.Button(pnl, label='Subject Allocation and Migration', pos=(250, 310), size=(200, 35))
        self.migrationButton.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.InitUI()

    def InitUI(self):
        self.rb1.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb2.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb3.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb4.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb5.Bind(wx.EVT_RADIOBUTTON, self.SetVal)
        self.rb6.Bind(wx.EVT_RADIOBUTTON, self.SetVal)

        self.confirmAdmissionButton.Bind(wx.EVT_BUTTON, self.OnConfirmAdmissionButtonPressed)
        self.stopAutoMigrationButton.Bind(wx.EVT_BUTTON, self.OnStopAutoMigrationButtonPressed)
        self.cancelAdmissionButton.Bind(wx.EVT_BUTTON, self.OnCancelAdmissionButton)
        self.migrationButton.Bind(wx.EVT_BUTTON, self.OnMigrationButtonPressed)

        self.SetSize((800, 600))
        self.Centre()

    def OnConfirmAdmissionButtonPressed(self, e):
        result, positions = process_admission(unit, "confirm")
        msg = unit + " unit admissions are confirmed successfully for merit position" + str(positions)
        wx.MessageBox(msg, 'Info', wx.OK | wx.ICON_INFORMATION)
        self.Close()

    def OnStopAutoMigrationButtonPressed(self, e):
        result, positions = process_admission(unit, "stop_migration")
        msg = unit + " unit stop auto migration requests are executed successfully for position " + str(positions)
        wx.MessageBox(msg, 'Info', wx.OK | wx.ICON_INFORMATION)
        self.Close()

    def OnCancelAdmissionButton(self, e):
        result, positions = process_admission(unit, "cancel")
        msg = unit + " unit admissions are cancelled successfully of position " + str(positions)
        wx.MessageBox(msg, 'Info', wx.OK | wx.ICON_INFORMATION)
        self.Close()

    def OnMigrationButtonPressed(self, e):
        result = execute_migration(unit)
        wx.MessageBox(result, 'Info', wx.OK | wx.ICON_INFORMATION)
        self.Close()

    def OnQuit(self, e):
        self.Close()

    def SetVal(self, e):
        global unit
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
        msg = "Unit " + unit + " is Selected"
        self.sb.SetStatusText(msg, 0)


def main():
    app = wx.App()
    ex = UI(None, title='JUST Migration Tool')
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
