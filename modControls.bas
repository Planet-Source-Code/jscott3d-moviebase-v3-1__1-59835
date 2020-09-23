Attribute VB_Name = "modControls"
Option Explicit
Dim i As Integer

'------------------------------------------------------------------
' Purpose   : Makes the status bar coding easer
'------------------------------------------------------------------
Public Sub StatusBarMsg(ByVal Msg As String, ByVal Panel As Integer)
    On Error Resume Next
    frmMain.StatusBar.Panels(Panel).Text = Msg 'Changes text in panels
End Sub

'---------------------------------------------------------------------------------
' Purpose   : Add Checked items
'---------------------------------------------------------------------------------
Public Sub AddListBoxItem(ByVal sItem As String, lst As ListBox)
    On Error Resume Next
    lst.AddItem sItem
End Sub
'---------------------------------------------------------------------------------
' Purpose   : Removes checked items
'---------------------------------------------------------------------------------
Public Sub RemoveListBoxItem(ByVal sItem As String, lst As ListBox)
    Dim iIdx As Integer
    On Error Resume Next
    For iIdx = 0 To lst.ListCount - 1
        If lst.List(iIdx) = sItem Then
            lst.RemoveItem iIdx
            Exit For
        End If
        DoEvents
    Next iIdx
End Sub
'---------------------------------------------------------------------------------
' Purpose   : Breaksdown listbox so can be stored in the database
'---------------------------------------------------------------------------------
Public Function lstBreakdown(lstTemp As ListBox)
    On Error Resume Next
    lstBroken = ""
    For i = 0 To lstTemp.ListCount - 1 'Loop thru the list box and save the data into a varible
        lstTemp.ListIndex = i
        If lstTemp.Text <> "" Then lstBroken = lstBroken & lstTemp.Text & "§" 'Creats the string seperated by §
    Next i
End Function
'---------------------------------------------------------------------------------
' Purpose   : Rebuilds Database string the loads into listbox
'---------------------------------------------------------------------------------
Public Function lstReBuild(lstTemp2 As ListBox, TempListCon As String)
    On Error Resume Next

    Dim vListBoxContents
    lstTemp2.Clear

    If InStr(1, TempListCon, "§") = 0 Then Exit Function

    'Fill The Variant Array
    vListBoxContents = Split(TempListCon, "§")

    'Loop Thru the Variant Array
    For i = 0 To UBound(vListBoxContents) - 1
        lstTemp2.AddItem vListBoxContents(i)
    Next i

End Function

'------------------------------------------------------------------
' Purpose   : Compacts Database and repares it
'------------------------------------------------------------------
Public Sub CompactDatabase(olddb As String, newdb As String, Optional locale, Optional options, Optional password)
    On Error Resume Next
    db.Close
    DBEngine.CompactDatabase olddb, newdb
    Kill DBPath
    Name newdb As DBPath

    LoadDataBase

End Sub

'---------------------------------------------------------------------------------
' Purpose   : Autocompetes Combo boxes
'---------------------------------------------------------------------------------

Public Sub ComboAutoComplete(ByRef SourceCtl As VB.ComboBox, _
    ByRef KeyAscii As Integer, ByRef LeftOffPos As Long)
  Dim iStart As Long
  Dim sSearchKey As String
  
  With SourceCtl
    'If text entered so far matches item(s) in the list, use autocomplete
    Select Case Chr$(KeyAscii)
      Case vbBack
        'Let backspace characters process as usual; otherwise try to match text
      Case Else
        If Chr$(KeyAscii) <> vbBack Then
          .SelText = Chr$(KeyAscii)
          
          iStart = .SelStart
          
          If LeftOffPos <> 0 Then
            .SelStart = LeftOffPos
            iStart = LeftOffPos
          End If
          
          sSearchKey = CStr(Left$(.Text, iStart))
          .ListIndex = SendMessage(.hWnd, CB_FINDSTRING, -1, _
              ByVal CStr(Left$(.Text, iStart)))
          
          If .ListIndex = -1 Then
            LeftOffPos = Len(sSearchKey)
          End If
          
          .SelStart = iStart
          .SelLength = Len(.Text)
          LeftOffPos = 0
          
          KeyAscii = 0
        End If
    End Select
  End With
End Sub

