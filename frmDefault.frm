VERSION 5.00
Begin VB.Form frmDefault 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Default"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDefault.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   221
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1740
      TabIndex        =   24
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4920
      Width           =   1455
   End
   Begin VB.ComboBox cboScreenRatio 
      Height          =   345
      Left            =   1020
      TabIndex        =   21
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Frame Frame8 
      Caption         =   "Disc Format"
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   3075
      Begin VB.OptionButton optFormat 
         Caption         =   "Dual-Sided"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optFormat 
         Caption         =   "Dual Layer"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optFormat 
         Caption         =   "Flipper"
         Height          =   195
         Index           =   3
         Left            =   1320
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optFormat 
         Caption         =   "Single Layer"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "NTSC / PAL"
      Height          =   675
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   3075
      Begin VB.OptionButton optNTSCPAL 
         Alignment       =   1  'Right Justify
         Caption         =   "NTSC"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   915
      End
      Begin VB.OptionButton optNTSCPAL 
         Caption         =   "PAL"
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   14
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Color"
      Height          =   675
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   3075
      Begin VB.OptionButton optColor 
         Caption         =   "Black/White"
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   12
         Top             =   240
         Width           =   1395
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Color"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.ComboBox cboNumberDisc 
      Height          =   345
      ItemData        =   "frmDefault.frx":058A
      Left            =   1500
      List            =   "frmDefault.frx":05AC
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox cboRegion 
      Height          =   345
      Left            =   1020
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.ComboBox cboLocation 
      Height          =   345
      Left            =   1020
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox cboPackaging 
      Height          =   345
      Left            =   1020
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.ComboBox cboType 
      Height          =   345
      Left            =   1020
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label22 
      Caption         =   "Screen Ratio:"
      Height          =   195
      Left            =   0
      TabIndex        =   22
      Top             =   1620
      Width           =   1575
   End
   Begin VB.Label Label16 
      Caption         =   "# of Discs (Tape):"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Region:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   "Location:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Packaging:"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label33 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force declaration of all variables
Dim strSearch As String

Private Sub cmdSave_Click()
    On Error Resume Next

    strSearch = "[Default] Like '" & "Default" & "'" ' compair Title to records

    With rsDefault
        .FindFirst strSearch 'find match to title
        .Edit
        !Type = cboType.Text
        !Packaging = cboPackaging.Text
        !Location = cboLocation.Text
        !Region = cboRegion.Text
        !NumofDisc = cboNumberDisc.Text
        !ScreenRatio = cboScreenRatio
        If optFormat.Item(0).Value = True Then !DiscFormat = 0
        If optFormat.Item(1).Value = True Then !DiscFormat = 1
        If optFormat.Item(2).Value = True Then !DiscFormat = 2
        If optFormat.Item(3).Value = True Then !DiscFormat = 3

        If optNTSCPAL.Item(0).Value = True Then !NTSCPAL = 0
        If optNTSCPAL.Item(1).Value = True Then !NTSCPAL = 1

        If optColor.Item(0).Value = True Then !Color = 0
        If optColor.Item(1).Value = True Then !Color = 1
        .Update
    End With
    Unload Me
End Sub

'------------------------------------------------------------------
' Public Controls
'------------------------------------------------------------------
'------------------------------------------------------------------
' Purpose   : Load CBOs
'------------------------------------------------------------------
Public Sub LoadCBO()
    On Error Resume Next
    rsType.MoveFirst
    Do While Not rsType.EOF
        cboType.AddItem rsType.Fields("Type")
        rsType.MoveNext
    Loop
    rsPackaging.MoveFirst
    Do While Not rsPackaging.EOF
        cboPackaging.AddItem rsPackaging.Fields("Packaging")
        rsPackaging.MoveNext
    Loop
    rsRegion.MoveFirst
    Do While Not rsRegion.EOF
        cboRegion.AddItem rsRegion.Fields("Region")
        rsRegion.MoveNext
    Loop
    rsLocation.MoveFirst
    Do While Not rsLocation.EOF
        cboLocation.AddItem rsLocation.Fields("Location")
        rsLocation.MoveNext
    Loop
    rsScreenRatio.MoveFirst
    Do While Not rsScreenRatio.EOF
        cboScreenRatio.AddItem rsScreenRatio.Fields("ScreenRatio")
        rsScreenRatio.MoveNext
    Loop

End Sub
'------------------------------------------------------------------
' Form Controls
'------------------------------------------------------------------
'--------------------------- Load
Private Sub Form_Load()
    On Error Resume Next
    LoadCBO
    With rsDefault
        .MoveFirst
        cboType.Text = .Fields("Type")
        cboPackaging.Text = .Fields("Packaging")
        cboLocation.Text = .Fields("Location")
        cboRegion.Text = .Fields("Region")
        cboNumberDisc.Text = .Fields("NumofDisc")
        cboScreenRatio = .Fields("ScreenRatio")
        optFormat(.Fields("DiscFormat")).Value = True
        optNTSCPAL(.Fields("NTSCPAL")).Value = True
        optColor(.Fields("Color")).Value = True
    End With

End Sub
'------------------------------------------------------------------
' Controls
'------------------------------------------------------------------
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
End Sub

