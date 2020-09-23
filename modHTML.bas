Attribute VB_Name = "modHTML"
Public txtHTML As String
Public ShowHTML As String
Public listHTML As String
Dim PageCol1, PageCol2, PageCol3 As String
Dim PageCol4, PageCol5, PageCol6 As String
Dim PageCol7, PageCol8, PageCol9 As String
Dim PageCol10, PageCol11, PageCol12 As String

Dim test As String

Dim i As Integer

Public Function CreatHTML(SearchText As String)
    Dim NTSCPAL As String
    Dim DiscTpy As String
    Dim Color As String
    Dim Review As String
    Dim Price As String

    strSearch = "[Title] Like '" & SearchText & "'"

    With rsMovies
        .FindFirst strSearch

        If .Fields("NTSCPAL") = 0 Then NTSCPAL = "NTSC"
        If .Fields("NTSCPAL") = 1 Then NTSCPAL = "PAL"

        If .Fields("DiscFormat") = 0 Then DiscTpy = "Dual Layer"
        If .Fields("DiscFormat") = 1 Then DiscTpy = "Single Layer"
        If .Fields("DiscFormat") = 2 Then DiscTpy = "Dual-Sided"
        If .Fields("DiscFormat") = 3 Then DiscTpy = "Flipper"

        If .Fields("Color") = 0 Then Color = "Color"
        If .Fields("Color") = 1 Then Color = "Black/White"

        If .Fields("UserReview") = "" Then

        ElseIf Left$(.Fields("UserReview"), 2) = 10 Then
            Review = Left$(.Fields("UserReview"), 2) & " / 10 ---" & Mid$(.Fields("UserReview"), 4)
        Else
            Review = Left$(.Fields("UserReview"), 1) & " / 10 ---" & Mid$(.Fields("UserReview"), 3)
        End If
        
        If .Fields("cost") = "" Then Price = "$0.00"
        If .Fields("cost") <> "" Then Price = "$" & .Fields("cost")
        

        txtHTML = ""
        txtHTML = txtHTML & "<html>" & vbCrLf
        txtHTML = txtHTML & "    <head>" & vbCrLf
        txtHTML = txtHTML & "        <title>MovieBase</title>" & vbCrLf
        txtHTML = txtHTML & "    </head>" & vbCrLf
        txtHTML = txtHTML & "    <body bgcolor=""#ffffff"">" & vbCrLf
        txtHTML = txtHTML & "        <div align=""center"">" & vbCrLf
        txtHTML = txtHTML & "            <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td colspan=""2"" bgcolor=""#e1e1e1"">" & vbCrLf
        txtHTML = txtHTML & "                        <div align=""left"">" & vbCrLf
        txtHTML = txtHTML & "                            <font size=""3"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Title") & "</font></div>" & vbCrLf
        txtHTML = txtHTML & "                    </td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#e1e1e1"" width=""12%""><font size=""3"" face=""Comic Sans MS"">" & vbCrLf & .Fields("MovieDate") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#e1e1e1"">" & vbCrLf
        txtHTML = txtHTML & "                        <div align=""right"">" & vbCrLf
        txtHTML = txtHTML & "                            <font size=""3"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Type") & "</font></div>" & vbCrLf
        txtHTML = txtHTML & "                    </td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td colspan=""4"">" & vbCrLf
        txtHTML = txtHTML & "                        <hr noshade>" & vbCrLf
        txtHTML = txtHTML & "                    </td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""15%""><font size=""2"" face=""Comic Sans MS"">Genre:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Genre") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""15%""><font size=""2"" face=""Comic Sans MS"">Sub Genre:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("SubGenre") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""10%""><font size=""2"" face=""Comic Sans MS"">Edition:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Edition") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""12%""><font size=""2"" face=""Comic Sans MS"">Director:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Director") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""10%""><font size=""2"" face=""Comic Sans MS"">Studio:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Studio") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""12%""><font size=""2"" face=""Comic Sans MS"">Series:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Series") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""10%""><font size=""2"" face=""Comic Sans MS"">Packaging:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Packaging") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""12%""><font size=""2"" face=""Comic Sans MS"">Location:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Location") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""10%""><font size=""2"" face=""Comic Sans MS"">Region:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Region") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""12%""><font size=""2"" face=""Comic Sans MS"">Rating:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Rating") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""10%""><font size=""2"" face=""Comic Sans MS"">User Review:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & Review & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""12%""><font size=""2"" face=""Comic Sans MS"">Length:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("Length") & " min.</font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""10%""><font size=""2"" face=""Comic Sans MS"">Date Purchased:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("DatePurched") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""12%""><font size=""2"" face=""Comic Sans MS"">Movie Date:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("MovieDate") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""10%""><font size=""2"" face=""Comic Sans MS"">DVD Date:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("DVDDate") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""12%""><font size=""2"" face=""Comic Sans MS""># of Discs:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("NumberDisc") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""10%""><font size=""2"" face=""Comic Sans MS"">Cost:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & Price & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""12%""></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td colspan=""4""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf
        txtHTML = txtHTML & "                            <hr>" & vbCrLf
        txtHTML = txtHTML & "                        </font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""10%""><font size=""2"" face=""Comic Sans MS"">Screen Ratio:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("ScreenRatio") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""12%""><font size=""2"" face=""Comic Sans MS"">Disc Format:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td>" & DiscTpy & "</td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""10%""><font size=""2"" face=""Comic Sans MS"">NTSC/PAL:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"">" & NTSCPAL & "</td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""12%""><font size=""2"" face=""Comic Sans MS"">Color:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"">" & Color & "</td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td colspan=""4"" ><font size=""2"" face=""Comic Sans MS"">" & vbCrLf
        txtHTML = txtHTML & "                            <hr>" & vbCrLf
        txtHTML = txtHTML & "                        </font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        Rebuild .Fields("SpecialFeatures")
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td valign=""top"" width=""10%""><font size=""2"" face=""Comic Sans MS"">Special Features:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td valign=""top"">" & vbCrLf
        txtHTML = txtHTML & "                        <dl>" & vbCrLf
        txtHTML = txtHTML & "                            <dt>" & ShowHTML
        txtHTML = txtHTML & "                        </dl>" & vbCrLf
        txtHTML = txtHTML & "                    </td>" & vbCrLf
        Rebuild .Fields("AudioTracks")
        txtHTML = txtHTML & "                    <td valign=""top"" width=""12%""><font size=""2"" face=""Comic Sans MS"">Audio Tracks:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td valign=""top"">" & vbCrLf
        txtHTML = txtHTML & "                        <dl>" & vbCrLf
        txtHTML = txtHTML & "                            <dt>" & ShowHTML
        txtHTML = txtHTML & "                        </dl>" & vbCrLf
        txtHTML = txtHTML & "                    </td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        Rebuild .Fields("Trailers")
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" valign=""top"" width=""10%""><font size=""2"" face=""Comic Sans MS"">Trailers:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" valign=""top"">" & vbCrLf
        txtHTML = txtHTML & "                        <dl>" & vbCrLf
        txtHTML = txtHTML & "                            <dt>" & ShowHTML
        txtHTML = txtHTML & "                        </dl>" & vbCrLf
        Rebuild .Fields("Subtitles")
        txtHTML = txtHTML & "                    </td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" valign=""top"" width=""12%""><font size=""2"" face=""Comic Sans MS"">Subtitles:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" valign=""top"">" & vbCrLf
        txtHTML = txtHTML & "                        <dl>" & vbCrLf
        txtHTML = txtHTML & "                            <dt>" & ShowHTML
        txtHTML = txtHTML & "                        </dl>" & vbCrLf
        txtHTML = txtHTML & "                    </td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td colspan=""4""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf
        txtHTML = txtHTML & "                            <hr>" & vbCrLf
        txtHTML = txtHTML & "                        </font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""10%""><font size=""2"" face=""Comic Sans MS""># of Movies:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("NumberMovies") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td width=""12%""><font size=""2"" face=""Comic Sans MS"">Free Time:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("FreeTime") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "                <tr>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""10%""><font size=""2"" face=""Comic Sans MS"">Mode:</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5""><font size=""2"" face=""Comic Sans MS"">" & vbCrLf & .Fields("TapeMode") & "</font></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5"" width=""12%""></td>" & vbCrLf
        txtHTML = txtHTML & "                    <td bgcolor=""#f5f5f5""></td>" & vbCrLf
        txtHTML = txtHTML & "                </tr>" & vbCrLf
        txtHTML = txtHTML & "            </table>" & vbCrLf
        txtHTML = txtHTML & "        </div>" & vbCrLf
        txtHTML = txtHTML & "    </body>" & vbCrLf
        txtHTML = txtHTML & "</html>" & vbCrLf


    End With

End Function
Public Function Rebuild(TempListCon As String)
    Dim vListBoxContents
    ShowHTML = ""
    If InStr(1, TempListCon, "§") = 0 Then Exit Function

    'Fill The Variant Array
    vListBoxContents = Split(TempListCon, "§")

    'Loop Thru the Variant Array
    For i = 0 To UBound(vListBoxContents) - 1
        ShowHTML = ShowHTML & "<dt><font size=""2"" face=""Comic Sans MS"">" & vListBoxContents(i) & "</font>"
    Next i

End Function

Public Function PrintList()
LoadHTMLList

listHTML = "<html>" & vbCrLf
listHTML = listHTML & "    <head>" & vbCrLf
listHTML = listHTML & "        <title>Welcome</title>" & vbCrLf
listHTML = listHTML & "    </head>" & vbCrLf
listHTML = listHTML & "    <body bgcolor=""#ffffff"">" & vbCrLf
listHTML = listHTML & "        <div align=""left"">" & vbCrLf

' First Page
listHTML = listHTML & "            <table width=""100%"" border=""0"" cellspacing=""2"" cellpadding=""0"">" & vbCrLf
listHTML = listHTML & "                <tr>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <div align=""left"">" & vbCrLf
listHTML = listHTML & "                            <dl>" & vbCrLf
listHTML = listHTML & "                                " & PageCol1 & vbCrLf
listHTML = listHTML & "                            </dl>" & vbCrLf
listHTML = listHTML & "                        </div>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <dl>" & vbCrLf
listHTML = listHTML & "                            " & PageCol2 & vbCrLf
listHTML = listHTML & "                        </dl>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <dl>" & vbCrLf
listHTML = listHTML & "                            " & PageCol3 & vbCrLf
listHTML = listHTML & "                        </dl>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                </tr>" & vbCrLf
listHTML = listHTML & "            </table>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
' Second Page
listHTML = listHTML & "            <table width=""100%"" border=""0"" cellspacing=""2"" cellpadding=""0"">" & vbCrLf
listHTML = listHTML & "                <tr>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <div align=""left"">" & vbCrLf
listHTML = listHTML & "                            <dl>" & vbCrLf
listHTML = listHTML & "                                " & PageCol4 & vbCrLf
listHTML = listHTML & "                            </dl>" & vbCrLf
listHTML = listHTML & "                        </div>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <dl>" & vbCrLf
listHTML = listHTML & "                            " & PageCol5 & vbCrLf
listHTML = listHTML & "                        </dl>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <dl>" & vbCrLf
listHTML = listHTML & "                            " & PageCol6 & vbCrLf
listHTML = listHTML & "                        </dl>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                </tr>" & vbCrLf
listHTML = listHTML & "            </table>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
' third Page
listHTML = listHTML & "            <table width=""100%"" border=""0"" cellspacing=""2"" cellpadding=""0"">" & vbCrLf
listHTML = listHTML & "                <tr>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <div align=""left"">" & vbCrLf
listHTML = listHTML & "                            <dl>" & vbCrLf
listHTML = listHTML & "                                " & PageCol7 & vbCrLf
listHTML = listHTML & "                            </dl>" & vbCrLf
listHTML = listHTML & "                        </div>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <dl>" & vbCrLf
listHTML = listHTML & "                            " & PageCol8 & vbCrLf
listHTML = listHTML & "                        </dl>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <dl>" & vbCrLf
listHTML = listHTML & "                            " & PageCol9 & vbCrLf
listHTML = listHTML & "                        </dl>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                </tr>" & vbCrLf
listHTML = listHTML & "            </table>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
listHTML = listHTML & "            <br>" & vbCrLf
' forth Page
listHTML = listHTML & "            <table width=""100%"" border=""0"" cellspacing=""2"" cellpadding=""0"">" & vbCrLf
listHTML = listHTML & "                <tr>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <div align=""left"">" & vbCrLf
listHTML = listHTML & "                            <dl>" & vbCrLf
listHTML = listHTML & "                                " & PageCol10 & vbCrLf
listHTML = listHTML & "                            </dl>" & vbCrLf
listHTML = listHTML & "                        </div>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <dl>" & vbCrLf
listHTML = listHTML & "                            " & PageCol11 & vbCrLf
listHTML = listHTML & "                        </dl>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                    <td align=""left"" valign=""top"" width=""33%"">" & vbCrLf
listHTML = listHTML & "                        <dl>" & vbCrLf
listHTML = listHTML & "                            " & PageCol12 & vbCrLf
listHTML = listHTML & "                        </dl>" & vbCrLf
listHTML = listHTML & "                    </td>" & vbCrLf
listHTML = listHTML & "                </tr>" & vbCrLf
listHTML = listHTML & "            </table>" & vbCrLf

listHTML = listHTML & "        </div>" & vbCrLf
listHTML = listHTML & "    </body>" & vbCrLf
listHTML = listHTML & "</html>"

End Function

Public Sub LoadHTMLList()
i = 1
With rsMovies
    .MoveFirst
    Do While Not .EOF
        Do While i <= 50
            If .EOF = True Then Exit Sub
            PageCol1 = PageCol1 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        Do While i <= 100
            If .EOF = True Then Exit Sub
            PageCol2 = PageCol2 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        Do While i <= 150
            If .EOF = True Then Exit Sub
            PageCol3 = PageCol3 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop

        Do While i <= 200
            If .EOF = True Then Exit Sub
            PageCol4 = PageCol4 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        Do While i <= 250
            If .EOF = True Then Exit Sub
            PageCol5 = PageCol5 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        Do While i <= 300
            If .EOF = True Then Exit Sub
            PageCol6 = PageCol6 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        Do While i <= 350
            If .EOF = True Then Exit Sub
            PageCol7 = PageCol7 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        Do While i <= 400
            If .EOF = True Then Exit Sub
            PageCol8 = PageCol8 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        Do While i <= 450
            If .EOF = True Then Exit Sub
            PageCol9 = PageCol9 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        Do While i <= 500
            If .EOF = True Then Exit Sub
            PageCol10 = PageCol10 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        Do While i <= 550
            If .EOF = True Then Exit Sub
            PageCol11 = PageCol11 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        Do While i <= 650
            If .EOF = True Then Exit Sub
            PageCol12 = PageCol12 & "<dt>" & i & " - " & .Fields("Title")
            i = i + 1
            .MoveNext
        Loop
        If .EOF = False Then .MoveNext
    Loop
End With
End Sub
