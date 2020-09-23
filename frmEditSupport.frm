VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditSupport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Support Table"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditSupport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.ComboBox cboTable 
         Height          =   345
         ItemData        =   "frmEditSupport.frx":058A
         Left            =   120
         List            =   "frmEditSupport.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1710
         Width           =   1575
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2460
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3210
         Width           =   1575
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Select a Table"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstDisplay 
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   7646
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   8819
      EndProperty
   End
End
Attribute VB_Name = "frmEditSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force declaration of all variables
Dim response As String

'------------------------------------------------------------------
' Public Controls
'------------------------------------------------------------------
Public Function LoadType(Optional Index As Integer)
    On Error Resume Next
    cboTable.ListIndex = intSelect
    Me.Caption = "Edit Table: " & cboTable.Text
    lstDisplay.ListItems.Clear
    Select Case intSelect
        Case 0
            rsGenre.MoveFirst
            Do While Not rsGenre.EOF
                lstDisplay.ListItems.Add , , rsGenre.Fields("Genre")
                rsGenre.MoveNext
            Loop
        Case 1
            rsEdition.MoveFirst
            Do While Not rsEdition.EOF
                lstDisplay.ListItems.Add , , rsEdition.Fields("Edition")
                rsEdition.MoveNext
            Loop
        Case 2
            rsStudio.MoveFirst
            Do While Not rsStudio.EOF
                lstDisplay.ListItems.Add , , rsStudio.Fields("Studio")
                rsStudio.MoveNext
            Loop
        Case 3
            rsPackaging.MoveFirst
            Do While Not rsPackaging.EOF
                lstDisplay.ListItems.Add , , rsPackaging.Fields("Packaging")
                rsPackaging.MoveNext
            Loop
        Case 4
            rsRegion.MoveFirst
            Do While Not rsRegion.EOF
                lstDisplay.ListItems.Add , , rsRegion.Fields("Region")
                rsRegion.MoveNext
            Loop
        Case 5
            rsRating.MoveFirst
            Do While Not rsRating.EOF
                lstDisplay.ListItems.Add , , rsRating.Fields("Ratings")
                rsRating.MoveNext
            Loop
        Case 6
            rsDirector.MoveFirst
            Do While Not rsDirector.EOF
                lstDisplay.ListItems.Add , , rsDirector.Fields("Director")
                rsDirector.MoveNext
            Loop
        Case 7
            rsSeries.MoveFirst
            Do While Not rsSeries.EOF
                lstDisplay.ListItems.Add , , rsSeries.Fields("Series")
                rsSeries.MoveNext
            Loop
        Case 8
            rsLocation.MoveFirst
            Do While Not rsLocation.EOF
                lstDisplay.ListItems.Add , , rsLocation.Fields("Location")
                rsLocation.MoveNext
            Loop
        Case 9
            rsType.MoveFirst
            Do While Not rsType.EOF
                lstDisplay.ListItems.Add , , rsType.Fields("Type")
                rsType.MoveNext
            Loop
        Case 10
            rsScreenRatio.MoveFirst
            Do While Not rsScreenRatio.EOF
                lstDisplay.ListItems.Add , , rsScreenRatio.Fields("ScreenRatio")
                rsScreenRatio.MoveNext
            Loop
        Case 11
            rsSpecialFeatures.MoveFirst
            Do While Not rsSpecialFeatures.EOF
                lstDisplay.ListItems.Add , , rsSpecialFeatures.Fields("SpecialFeature")
                rsSpecialFeatures.MoveNext
            Loop
        Case 12
            rsTrailers.MoveFirst
            Do While Not rsTrailers.EOF
                lstDisplay.ListItems.Add , , rsTrailers.Fields("Trailers")
                rsTrailers.MoveNext
            Loop
        Case 13
            rsAudio.MoveFirst
            Do While Not rsAudio.EOF
                lstDisplay.ListItems.Add , , rsAudio.Fields("Audio")
                rsAudio.MoveNext
            Loop
        Case 14
            rsSubtitles.MoveFirst
            Do While Not rsSubtitles.EOF
                lstDisplay.ListItems.Add , , rsSubtitles.Fields("Subtitles")
                rsSubtitles.MoveNext
            Loop
    End Select

End Function
'------------------------------------------------------------------
' Form Controls
'------------------------------------------------------------------
'--------------------------- Load
Private Sub Form_Load()
    Dim i As Integer
    On Error Resume Next

    i = 0

    rsCBO.MoveFirst
    Do While Not rsCBO.EOF
        cboTable.AddItem rsCBO.Fields("CBOName")
        rsCBO.MoveNext
    Loop

    LoadType intSelect


End Sub
'------------------------------------------------------------------
' Controls
'------------------------------------------------------------------
'---------- CBOs
'----- Table Type
Private Sub cboTable_Click()
    On Error Resume Next
    intSelect = cboTable.ListIndex 'changes type
    LoadType intSelect 'reloads list
End Sub
'---------- Command Buttons
'----- CLose
Private Sub cmdClose_Click()
    On Error Resume Next
    Unload Me
End Sub
'----- Refresh
Private Sub cmdRefresh_Click()
    On Error Resume Next
    LoadType intSelect
End Sub
'----- Clear
Private Sub cmdClear_Click()
    On Error Resume Next
    lstDisplay.ListItems.Clear
End Sub
'----- Add Button
Private Sub cmdAdd_Click()
    On Error Resume Next
    response = InputBox("Add a new selection", "Input it needed")
    Select Case intSelect
        Case 0
            With rsGenre
                .AddNew
                !Genre = response
                .Update
            End With
        Case 1
            With rsEdition
                .AddNew
                !Edition = response
                .Update
            End With
        Case 2
            With rsStudio
                .AddNew
                !Studio = response
                .Update
            End With
        Case 3
            With rsPackaging
                .AddNew
                !Packaging = response
                .Update
            End With
        Case 4
            With rsRegion
                .AddNew
                !Region = response
                .Update
            End With
        Case 5
            With rsRating
                .AddNew
                !Rating = response
                .Update
            End With
        Case 6
            With rsDirector
                .AddNew
                !Director = response
                .Update
            End With
        Case 7
            With rsSeries
                .AddNew
                !Series = response
                .Update
            End With
        Case 8
            With rsLocation
                .AddNew
                !Location = response
                .Update
            End With
        Case 9
            With rsType
                .AddNew
                !Type = response
                .Update
            End With
        Case 10
            With rsScreenRatio
                .AddNew
                !ScreenRatio = response
                .Update
            End With
        Case 11
            With rsSpecialFeatures
                .AddNew
                !SpecialFeatures = response
                .Update
            End With
        Case 12
            With rsTrailers
                .AddNew
                !Trailers = response
                .Update
            End With
        Case 13
            With rsAudio
                .AddNew
                !Audio = response
                .Update
            End With
        Case 14
            With rsSubtitles
                .AddNew
                !Subtitles = response
                .Update
            End With
    End Select
    LoadType intSelect

    'frmAddEdit.ClearCBO
End Sub
'----- Delete Button
Private Sub cmdDelete_Click()
    Dim SearchText, strSearch As String
    On Error Resume Next

    SearchText = lstDisplay.SelectedItem.Text '// set the search text as the selected in the listbox
    response = MsgBox("Are You Sure That You Want To Delete This? :: " & SearchText, vbYesNo)   '// makes sure that you want to delete the selected

    If response = vbYes Then ' User chose Yes.
        Select Case intSelect
            Case 0
                strSearch = "[Genre] Like '" & SearchText & "'" '// compares to selected'
                With rsGenre '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 1
                strSearch = "[Edition] Like '" & SearchText & "'" '// compares to selected'
                With rsEdition '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 2
                strSearch = "[Studio] Like '" & SearchText & "'" '// compares to selected'
                With rsStudio '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 3
                strSearch = "[Packaging] Like '" & SearchText & "'" '// compares to selected'
                With rsPackaging '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 4
                strSearch = "[Region] Like '" & SearchText & "'" '// compares to selected'
                With rsRegion '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 5
                strSearch = "[Rating] Like '" & SearchText & "'" '// compares to selected'
                With rsRating '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 6
                strSearch = "[Director] Like '" & SearchText & "'" '// compares to selected'
                With rsDirector '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 7
                strSearch = "[Series] Like '" & SearchText & "'" '// compares to selected'
                With rsSeries '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 8
                strSearch = "[Location] Like '" & SearchText & "'" '// compares to selected'
                With rsLocation '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 9
                strSearch = "[Type] Like '" & SearchText & "'" '// compares to selected'
                With rsType '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 10
                strSearch = "[ScreenRatio] Like '" & SearchText & "'" '// compares to selected'
                With rsScreenRatio '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 11
                strSearch = "[SpecialFeature] Like '" & SearchText & "'" '// compares to selected'
                With rsSpecialFeatures '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 12
                strSearch = "[Trailers] Like '" & SearchText & "'" '// compares to selected'
                With rsTrailers '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 13
                strSearch = "[Audio] Like '" & SearchText & "'" '// compares to selected'
                With rsAudio '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
            Case 14
                strSearch = "[Subtitles] Like '" & SearchText & "'" '// compares to selected'
                With rsSubtitles '// opens the recordset
                    .FindFirst strSearch '// looks for selected in recordset
                    .Delete '// deletes the selected
                End With '// closed the recordset
        End Select
    Else
        Exit Sub
    End If
    LoadType intSelect

    'frmAddEdit.ClearCBO

End Sub
