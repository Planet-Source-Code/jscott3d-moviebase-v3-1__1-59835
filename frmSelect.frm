VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   231
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "None"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin MSComctlLib.ListView lstSelect 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
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
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force declaration of all variables
Dim iIdx, iIdx2 As Integer

'------------------------------------------------------------------
' Public Controls
'------------------------------------------------------------------
'----- Gets Checked from parent form
Public Function GetChecked(lst As ListBox)
    Dim temp, temp2 As String
    On Error Resume Next


    For iIdx = 0 To lstSelect.ListItems.Count - 1
        temp = lstSelect.ListItems.Item(iIdx + 1)

        For iIdx2 = 0 To lst.ListCount - 1
            temp2 = lst.List(iIdx2)

            If temp = temp2 Then lstSelect.ListItems.Item(iIdx + 1).Checked = True

        Next iIdx2
        DoEvents
    Next iIdx

End Function
'-----loads listbox
Public Sub LoadList()
    On Error Resume Next
    Select Case intSelect
        Case 11
            rsSpecialFeatures.MoveFirst
            Do While Not rsSpecialFeatures.EOF
                lstSelect.ListItems.Add , , rsSpecialFeatures.Fields("SpecialFeature")
                rsSpecialFeatures.MoveNext
            Loop
        Case 12
            rsTrailers.MoveFirst
            Do While Not rsTrailers.EOF
                lstSelect.ListItems.Add , , rsTrailers.Fields("Trailers")
                rsTrailers.MoveNext
            Loop
        Case 13
            rsAudio.MoveFirst
            Do While Not rsAudio.EOF
                lstSelect.ListItems.Add , , rsAudio.Fields("Audio")
                rsAudio.MoveNext
            Loop
        Case 14
            rsSubtitles.MoveFirst
            Do While Not rsSubtitles.EOF
                lstSelect.ListItems.Add , , rsSubtitles.Fields("Subtitles")
                rsSubtitles.MoveNext
            Loop

    End Select

End Sub

Private Sub cmdAll_Click()
    Dim i As Integer
    i = 1
    Do While i <= lstSelect.ListItems.Count
        lstSelect.ListItems.Item(i).Checked = True
        Select Case intSelect
            Case 11
                If lstSelect.ListItems.Item(i).Checked = True Then AddListBoxItem (lstSelect.ListItems.Item(i).Text), frmAddEdit.lstSpecialFeatures
            Case 12
                If lstSelect.ListItems.Item(i).Checked = True Then AddListBoxItem (lstSelect.ListItems.Item(i).Text), frmAddEdit.lstTrailers
            Case 13
                If lstSelect.ListItems.Item(i).Checked = True Then AddListBoxItem (lstSelect.ListItems.Item(i).Text), frmAddEdit.lstAudioTracks
            Case 14
                If lstSelect.ListItems.Item(i).Checked = True Then AddListBoxItem (lstSelect.ListItems.Item(i).Text), frmAddEdit.lstSubtitles
        End Select
        i = i + 1
    Loop
End Sub


Private Sub cmdNone_Click()
    Dim i As Integer
    i = 1
    Do While i <= lstSelect.ListItems.Count
        lstSelect.ListItems.Item(i).Checked = False
        Select Case intSelect
            Case 11
                If lstSelect.ListItems.Item(i).Checked = False Then RemoveListBoxItem (lstSelect.ListItems.Item(i).Text), frmAddEdit.lstSpecialFeatures
            Case 12
                If lstSelect.ListItems.Item(i).Checked = False Then RemoveListBoxItem (lstSelect.ListItems.Item(i).Text), frmAddEdit.lstTrailers
            Case 13
                If lstSelect.ListItems.Item(i).Checked = False Then RemoveListBoxItem (lstSelect.ListItems.Item(i).Text), frmAddEdit.lstAudioTracks
            Case 14
                If lstSelect.ListItems.Item(i).Checked = False Then RemoveListBoxItem (lstSelect.ListItems.Item(i).Text), frmAddEdit.lstSubtitles

        End Select
        i = i + 1
    Loop

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
    LoadList

    lstSelect.ColumnHeaders.Item(1).Width = lstSelect.Width - 5

End Sub
'------------------------------------------------------------------
' Controls
'------------------------------------------------------------------
'---------- Command Button
'----- Ok Button
Private Sub cmdOK_Click()
    On Error Resume Next
    Unload Me
End Sub
'----- List box : checks to parent form
Private Sub lstSelect_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Select Case intSelect
        Case 11
            If Item.Checked = True Then
                AddListBoxItem (Item.Text), frmAddEdit.lstSpecialFeatures
            Else
                RemoveListBoxItem (Item.Text), frmAddEdit.lstSpecialFeatures
            End If
        Case 12
            If Item.Checked = True Then
                AddListBoxItem (Item.Text), frmAddEdit.lstTrailers
            Else
                RemoveListBoxItem (Item.Text), frmAddEdit.lstTrailers
            End If
        Case 13
            If Item.Checked = True Then
                AddListBoxItem (Item.Text), frmAddEdit.lstAudioTracks
            Else
                RemoveListBoxItem (Item.Text), frmAddEdit.lstAudioTracks
            End If
        Case 14
            If Item.Checked = True Then
                AddListBoxItem (Item.Text), frmAddEdit.lstSubtitles
            Else
                RemoveListBoxItem (Item.Text), frmAddEdit.lstSubtitles
            End If
    End Select


End Sub


