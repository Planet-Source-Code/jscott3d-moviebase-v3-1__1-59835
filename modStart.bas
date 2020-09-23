Attribute VB_Name = "modStart"
Option Explicit 'force declaration of all variables

'---------------------------------------------------------------------------------
' Purpose   : Locates the MovieBase Database
'---------------------------------------------------------------------------------
Public Function DBPath() As String
    On Error Resume Next
    DBPath = GetSetting(App.EXEName, "Settings", "Path") 'looks for settings

    If DBPath = "" Then 'settings empty
        DBPath = App.Path & "\Database\MovieBase.mdb" 'Database Path (by default)
        SaveSetting App.EXEName & Version, "Settings", "Path", DBPath 'Saves db path
    End If

End Function

'---------------------------------------------------------------------------------
' Purpose   : Main Loader for app
'---------------------------------------------------------------------------------
Public Sub Main()
    Dim RunOnce As String, DateRan As String
    On Error Resume Next

    Version = " " & App.Major & "." & App.Minor & "." & App.Revision

    RunOnce = GetSetting(App.EXEName & Version, "RunOnce", "RunOnce") 'setting if Ran
    DateRan = Date 'First run Date

    Select Case RunOnce 'select if ran or not
        Case 1 'Has been ran
            LoadDataBase
            LoadDataVaris
            frmMain.Show 'Loads main form
        Case Else 'First Run
            SaveSetting App.EXEName & Version, "RunOnce", "RunOnce", "1" 'Changes to has been ran
            SaveSetting App.EXEName & Version, "RunOnce", "DateRan", DateRan 'First Run Date
            SaveSetting App.EXEName & Version, "Settings", "Path", DBPath 'Saves db path
            SaveSetting App.EXEName & Version, "Settings", "DefaultStart", 1
            SaveSetting App.EXEName & Version, "Settings", "SizerBarLR", 280
            SaveSetting App.EXEName & Version, "Settings", "SizerBarTB", 451

            LoadDataBase
            LoadDataVaris
            frmMain.Show 'Loads main form
    End Select
End Sub

