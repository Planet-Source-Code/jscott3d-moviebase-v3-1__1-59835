VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDatabaseChange 
      Caption         =   "Change"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   75
      Width           =   855
   End
   Begin VB.TextBox txtDatabaseLocation 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   4455
   End
   Begin VB.CheckBox chkDefaults 
      Caption         =   "(on/off) Load defaults on new Movie."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Database Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force declaration of all variables

'------------------------------------------------------------------
' Public Controls
'------------------------------------------------------------------
Public Sub LoadSettings()
    On Error Resume Next
    chkDefaults.Value = DefaultStartup
    txtDatabaseLocation.Text = DBPath
End Sub

Private Sub cmdDatabaseChange_Click()
    On Error GoTo ErrHandler

    With CommonDialog1
        .CancelError = True
        .InitDir = App.Path & "\Database"
        '.Flags = cd10FNHideReadOnly
        .Filter = "Database Files (*.mdb)|*.mdb"
        .FilterIndex = 1
        .ShowOpen
    End With

    txtDatabaseLocation.Text = CommonDialog1.FileName

ErrHandler:
    Exit Sub

End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    SaveSetting App.EXEName, "Settings", "DefaultStart", chkDefaults.Value
    SaveSetting App.EXEName, "Settings", "Path", txtDatabaseLocation.Text

    LoadDataVaris
    Unload Me
End Sub

'------------------------------------------------------------------
' Purpose   :
'------------------------------------------------------------------
'------------------------------------------------------------------
' Form Controls
'------------------------------------------------------------------
'--------------------------- Load
Private Sub Form_Load()
    On Error Resume Next
    LoadSettings
End Sub
'------------------------------------------------------------------
' Controls
'------------------------------------------------------------------
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
End Sub
