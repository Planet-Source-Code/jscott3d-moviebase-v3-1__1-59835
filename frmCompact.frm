VERSION 5.00
Begin VB.Form frmCompact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compact and/or Repair The Database"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCompactdba 
      Caption         =   "Compact Database"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmCompact.frx":058A
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label lblNewSize 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label lblFreeSpace 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5055
   End
End
Attribute VB_Name = "frmCompact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force declaration of all variables
Dim dbasize As Long
Dim response As String

'------------------------------------------------------------------
' Form Controls
'------------------------------------------------------------------
'--------------------------- Load
Private Sub Form_Load()
    Dim fs, d, s
    Dim drvpath As String
    Dim freeSpace As Long

    drvpath = App.Path
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvpath))

    freeSpace = d.AvailableSpace / 1024 / 1024
    s = "Drive " & Left(App.Path, 1) & " has "
    lblFreeSpace = s & FormatNumber(freeSpace, 0) & "MB free"

    dbasize = FileLen(DBPath)
    lblSize = "Current Database size: " & Format((dbasize / 1024) / 1024, "standard") & "MB."


    On Error GoTo err
    If freeSpace * 1024 * 1024 < dbasize Then
        lblNewSize = "Not enough space to compact database clear some space on drive " & Left(App.Path, 1)
        cmdCompactdba.Enabled = False
    End If
err:
    Exit Sub

End Sub
'------------------------------------------------------------------
' Controls
'------------------------------------------------------------------
'----- Close Button
Private Sub cmdClose_Click()
    On Error Resume Next
    Unload Me
End Sub
'----- Compact repair button
Private Sub cmdCompactdba_Click()
    On Error Resume Next
    CompactDatabase DBPath, App.Path & "\Database\dbase.bak"

    dbasize = FileLen(DBPath)
    lblNewSize = "Compacted Database size : " & Format((dbasize / 1024) / 1024, "standard") & "MB."

    cmdClose.Visible = True
    cmdCompactdba.Enabled = False

End Sub

