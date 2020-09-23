Attribute VB_Name = "modDatabase"
Option Explicit 'force declaration of all variables

Public db As Database 'Database
Public rsMovies As Recordset 'Movies table
Public rsTreeType As Recordset  'Movie tree Type
Public rsGenre  As Recordset ' Movie Genres
Public rsRating  As Recordset ' Movie Ratings
Public rsRegion  As Recordset ' Movie Ratings
Public rsType  As Recordset ' Types of Movie
Public rsStudio  As Recordset ' Studio
Public rsEdition  As Recordset ' Edition
Public rsDirector  As Recordset ' Director
Public rsPackaging  As Recordset ' Packaging
Public rsSeries  As Recordset ' Series
Public rsLocation  As Recordset ' Location
Public rsScreenRatio  As Recordset ' RAtio
Public rsAudio  As Recordset ' Audio Tracks
Public rsSubtitles  As Recordset ' Subtitles
Public rsTrailers As Recordset ' Movie Trailers
Public rsSpecialFeatures As Recordset ' Special Features
Public rsCBO  As Recordset ' Edit CBOs
Public rsDefault As Recordset 'Default Settings of AddMovie
Public rsSearchType As Recordset 'Default Search types
'---------------------------------------------------------------------------------
' Purpose   :Loads all dabase tables
'---------------------------------------------------------------------------------
Public Sub LoadDataBase()
    On Error Resume Next

    Set db = OpenDatabase(DBPath) 'opendatabase
    Set rsMovies = db.OpenRecordset("tblMovies", dbOpenDynaset)  'open movie table
    Set rsTreeType = db.OpenRecordset("tblTreeType", dbOpenDynaset)
    Set rsGenre = db.OpenRecordset("tblGenre", dbOpenDynaset)
    Set rsRating = db.OpenRecordset("tblRatings", dbOpenDynaset)
    Set rsRegion = db.OpenRecordset("tblRegion", dbOpenDynaset)
    Set rsType = db.OpenRecordset("tblType", dbOpenDynaset)
    Set rsStudio = db.OpenRecordset("tblStudio", dbOpenDynaset)
    Set rsEdition = db.OpenRecordset("tblEdition", dbOpenDynaset)
    Set rsDirector = db.OpenRecordset("tblDirector", dbOpenDynaset)
    Set rsPackaging = db.OpenRecordset("tblPackaging", dbOpenDynaset)
    Set rsSeries = db.OpenRecordset("tblSeries", dbOpenDynaset)
    Set rsLocation = db.OpenRecordset("tblLocation", dbOpenDynaset)
    Set rsScreenRatio = db.OpenRecordset("tblScreenRatio", dbOpenDynaset)
    Set rsAudio = db.OpenRecordset("tblAudio", dbOpenDynaset)
    Set rsSubtitles = db.OpenRecordset("tblSubtitles", dbOpenDynaset)
    Set rsTrailers = db.OpenRecordset("tblTrailers", dbOpenDynaset)
    Set rsSpecialFeatures = db.OpenRecordset("tblSpecialFeatures", dbOpenDynaset)
    Set rsCBO = db.OpenRecordset("tblCBO", dbOpenDynaset)
    Set rsDefault = db.OpenRecordset("tblDefault", dbOpenDynaset)
    Set rsSearchType = db.OpenRecordset("tblSearchType", dbOpenDynaset)

End Sub

'---------------------------------------------------------------------------------
' Purpose   : Loads frmMain Treeview
'---------------------------------------------------------------------------------
Public Function fillDVDTreeView(treView As TreeView, strSortKey As String)
    On Error Resume Next
    Dim strTemp As String
    Dim nodeX As Node
    Dim nodeX2 As Node
    Dim i As Integer

    rsMovies.MoveFirst 'move to first record
    Do While Not rsMovies.EOF 'loop until end of records
        frmMain.CoolBar1.Bands.Item(3).Caption = "Total Movies: " & rsMovies.RecordCount  'displays number of movies by record count
        rsMovies.MoveNext 'moves to next record
    Loop

    treView.Visible = False 'hide treeview for faster load

    treView.Nodes.Clear 'clear tree
    treView.Nodes.Add , , "Collection", "MovieBase.mdb", "Open" ' add top node, collection as parent node
    treView.Nodes.Item(1).Expanded = True 'expand tree at first node

    Select Case strSortKey
        Case "Title"    'sort by title
            rsMovies.MoveFirst 'move to first record
            Do While Not rsMovies.EOF 'loop until end of records
                Set nodeX = treView.Nodes.Add("Collection", tvwChild, , rsMovies.Fields("Title"), "File")  'add node with titles
                nodeX.Tag = "T|" & rsMovies.Fields("MovieID")
                rsMovies.MoveNext 'moves to next record
            Loop
            GoTo Finished ' finished with loading the treeview

        Case "Genre"
            rsGenre.MoveFirst
            Do While Not rsGenre.EOF
                Set nodeX2 = treView.Nodes.Add("Collection", tvwChild, rsGenre.Fields("Genre"), rsGenre.Fields("Genre"), "Closed")
                nodeX2.Tag = "G|" & rsGenre.Fields("Genre")
                rsGenre.MoveNext
            Loop
            rsMovies.MoveFirst
            Do While Not rsMovies.EOF
                strTemp = rsMovies.Fields("Genre")
                Set nodeX = treView.Nodes.Add(strTemp, tvwChild, , rsMovies.Fields("Title"), "File")
                nodeX.Tag = "T|" & rsMovies.Fields("MovieID")
                rsMovies.MoveNext
            Loop
            GoTo Finished


        Case "Rating"
            rsRating.MoveFirst
            Do While Not rsRating.EOF
                Set nodeX2 = treView.Nodes.Add("Collection", tvwChild, rsRating.Fields("Ratings"), rsRating.Fields("Ratings"), "Closed")
                nodeX2.Tag = "G|" & rsRating.Fields("Ratings")
                rsRating.MoveNext
            Loop
            rsMovies.MoveFirst
            Do While Not rsMovies.EOF
                strTemp = rsMovies.Fields("Rating")
                Set nodeX = treView.Nodes.Add(strTemp, tvwChild, , rsMovies.Fields("Title"), "File")
                nodeX.Tag = "T|" & rsMovies.Fields("MovieID")
                rsMovies.MoveNext
            Loop
            GoTo Finished

        Case "Region"
            rsRegion.MoveFirst
            Do While Not rsRegion.EOF
                Set nodeX2 = treView.Nodes.Add("Collection", tvwChild, rsRegion.Fields("Region"), rsRegion.Fields("Region"), "Closed")
                nodeX2.Tag = "G|" & rsRegion.Fields("Region")
                rsRegion.MoveNext
            Loop
            rsMovies.MoveFirst
            Do While Not rsMovies.EOF
                strTemp = rsMovies.Fields("Region")
                Set nodeX = treView.Nodes.Add(strTemp, tvwChild, , rsMovies.Fields("Title"), "File")
                nodeX.Tag = "T|" & rsMovies.Fields("MovieID")
                rsMovies.MoveNext
            Loop
            GoTo Finished
        Case "Format"
            rsType.MoveFirst
            Do While Not rsType.EOF
                Set nodeX2 = treView.Nodes.Add("Collection", tvwChild, rsType.Fields("Type"), rsType.Fields("Type"), "Closed")
                nodeX2.Tag = "G|" & rsType.Fields("Type")
                rsType.MoveNext
            Loop
            rsMovies.MoveFirst
            Do While Not rsMovies.EOF
                strTemp = rsMovies.Fields("Type")
                Set nodeX = treView.Nodes.Add(strTemp, tvwChild, , rsMovies.Fields("Title"), "File")
                nodeX.Tag = "T|" & rsMovies.Fields("MovieID")
                rsMovies.MoveNext
            Loop
            GoTo Finished

    End Select

Finished:

    i = 2
    Do Until i = treView.Nodes.Count
        treView.Nodes.Item(i).Sorted = True
        i = i + 1
    Loop
    If strSortKey = "Title" Then treView.Nodes.Item(1).Sorted = True

    treView.Visible = True 'Shows the treeview after filling
End Function
