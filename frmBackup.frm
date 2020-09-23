VERSION 5.00
Begin VB.Form frmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup The Database"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDestination 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   5295
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Backup Database"
      Height          =   255
      Left            =   1980
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
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
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label lblDbaSize 
      Caption         =   "Current Database Size is:"
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
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Backup Destination"
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
      TabIndex        =   3
      Top             =   480
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbasize As Long
Dim response As String

'------------------------------------------------------------------
' Form Controls
'------------------------------------------------------------------
'--------------------------- Load
Private Sub Form_Load()
    On Error Resume Next
    dbasize = FileLen(DBPath) 'determines db size
    lblSize = Format((dbasize / 1024) / 1024, "standard") & "MB." ' displays db size
    txtDestination.Text = App.Path & "\Database\Backup - " & Format(Date, "yyyy.mm.dd") ' sets backup path

End Sub

'------------------------------------------------------------------
' Controls
'------------------------------------------------------------------
'----- Backup Button
Private Sub cmdBackup_Click()
    On Error Resume Next
    If txtDestination <> "" Then
        db.Close ' Closes database

        Dim objFSO As New FileSystemObject 'new object
        Dim objFile As File 'new file

        Set objFile = objFSO.GetFile(DBPath) 'set file as db
        ' if folder exists then copy else creat folder then copy
        If objFSO.FolderExists(txtDestination.Text) = True Then
            objFile.Copy txtDestination.Text & "\MovieBase.mdb"
        Else
            MkDir txtDestination.Text 'creats folder
            objFile.Copy txtDestination.Text & "\MovieBase.mdb"
        End If

        LoadDataBase 'reload database
        GoTo Done

    ElseIf txtDestination = "" Then
        MsgBox "You must specify a distination for the backup", vbCritical
    End If

    Exit Sub
Done:
    lblStatus.Caption = "Backup Complete"
    If MsgBox("Backup Complete.  Click OK to exit.") = vbOK Then Unload Me
End Sub
