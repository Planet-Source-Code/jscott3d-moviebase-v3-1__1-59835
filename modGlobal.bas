Attribute VB_Name = "modGlobal"
Option Explicit

Public Version As String

Public Const CONTROL_PADDING = 0 'spacing around controls
Public Const CONTROL_PADDING_BAR = 5 'sizer bar Width/Height
Public iSizerPosLR As Integer 'position of the sizer Left/Right
Public iSizerPosTB As Integer 'position of the sizer Top/Bottom
Public TPPX As Integer, TPPY As Integer 'shortcuts for TwipsPerPixel


Public Site As String 'Site address for winstock
Public Port As Integer 'port for winstock
Public Edit As Boolean ' between add and edit on the form
Public EditID As String ' between add and edit on the form
Public lstBroken As String 'Used for breakdown of listboxes
Public lstFixed As String ' used for the rebuild of listboxes

Public DefaultStartup As Integer 'settings for Defaults in add mode
Public BarLR As Integer
Public BarTB As Integer
Public intSelect As Integer '// Determines Type of Select or edit support
        '0= Genre, 1 = Edition, 2 = Studio, 3 = Packaging, 4= Region,
        '5 = Rating, 6 = Director, 7 = Series, 8 = Location, 9= Type,
        '10 = Screen Ratio, 11 = Special Features, 12 = Trailers,
        '13 = Audio Tracks, 14 = Subtitles

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) _
    As Long

Public Const CB_FINDSTRING As Long = &H14C

'---------------------------------------------------------------------------------
' Purpose   : Loads all varis
'---------------------------------------------------------------------------------
Public Sub LoadDataVaris()
    On Error Resume Next
    Site = "www.imdb.com" 'Site Address
    Port = 80 'Port number

    DefaultStartup = GetSetting(App.EXEName & Version, "Settings", "DefaultStart")
    
    BarLR = GetSetting(App.EXEName & Version, "Settings", "SizerBarLR")
    BarTB = GetSetting(App.EXEName & Version, "Settings", "SizerBarTB")
End Sub

'---------------------------------------------------------------------------------
' Purpose   : Save all varis
'---------------------------------------------------------------------------------
Public Sub SaveDataVaris()
    SaveSetting App.EXEName & Version, "Settings", "SizerBarLR", frmMain.picResizeLR.Left
    SaveSetting App.EXEName & Version, "Settings", "SizerBarTB", frmMain.picResizeTB.Top

End Sub



