VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRestore 
   Caption         =   "Restore The Database"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   Icon            =   "frmRestore.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   158
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore Database"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2205
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5535
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Restore from"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Current Database Size:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblSize 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblSelectedDba 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1560
      Width           =   5895
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force declaration of all variables
Dim dbasize As Long
Dim dbasize2 As Long
Dim response As String

'------------------------------------------------------------------
' Form Controls
'------------------------------------------------------------------
'--------------------------- Load
Private Sub Form_Load()
    dbasize = FileLen(DBPath) 'database size
    lblSize = Format((dbasize / 1024) / 1024, "standard") & "MB." 'displays db size

End Sub
'------------------------------------------------------------------
' Controls
'------------------------------------------------------------------
'----- Source button
Private Sub cmdSource_Click()
    On Error GoTo ErrHandler

    With CommonDialog1 'creats an open dialog for the files
        .CancelError = True
        .Filter = "Database Files (*.mdb)|*.mdb"
        .FilterIndex = 1
        .ShowOpen
    End With

    txtSource.Text = CommonDialog1.FileName 'displays filename
    dbasize2 = FileLen(txtSource) 'gets size of new db
    lblSelectedDba = "Selected Backup Database is : " & Format((dbasize2 / 1024) / 1024, "standard") & "MB." 'displays size of new db
    If Right$(txtSource.Text, 4) = ".mdb" Then cmdRestore.Enabled = True 'if file is a database then enable restore button

    Exit Sub
ErrHandler:
    MsgBox err.Description, vbCritical, err.Number
End Sub
'----- Restore button
Private Sub cmdRestore_Click()
    On Error Resume Next
    If MsgBox("Restoring database from location " & txtSource.Text & " will replace existing database files. Do you want to Contunue", vbYesNo) = vbYes Then
        db.Close
        Dim objFSO As New FileSystemObject
        Dim objFile As File

        Set objFile = objFSO.GetFile(txtSource.Text)
        objFile.Copy DBPath
        GoTo Done
    Else
        lblStatus.Caption = "Database Restore Canceled"
    End If
    Exit Sub
Done:
    lblStatus = "Restore Complete"
    response = MsgBox("Database Restored Click Ok to close")
    LoadDataBase
    fillDVDTreeView frmMain.treMovieList, "Title" 'loads treview sorted by title

    If response = True Then Unload Me

End Sub

