Attribute VB_Name = "DatabaseFunctions"
Option Explicit

Public HN As Hini, CurrPlayList As String
Public Sub NewPlaylist(filename As String)
    Dim tempstr As String
    If HN.existsancekey("Playlists", filename) = False Then
        tempstr = Right(filename, Len(filename) - InStrRev(filename, "\"))
        If InStr(tempstr, ".") > 0 Then tempstr = Left(tempstr, InStrRev(tempstr, ".") - 1)
        HN.createkey "Playlists", filename, tempstr
    End If
End Sub
Public Function GetPlaylist(index As Long) As String
    GetPlaylist = HN.keyindex(HN.qualifiedsectionhandle("Playlists"), index, 1)
End Function
Public Function ScanFile(filename As String) As Boolean
    Dim temp As String
    temp = Replace(filename, "\", "|")
    If HN.existsancekey("Songs\" & temp, "Date") = True Then
        temp = HN.getkeycontents("Songs\" & temp, "Date")
        If HasBeenModified(filename, temp) Then
            DeleteFileDetails filename
            AddFileDetails filename
        End If
    Else
        AddFileDetails filename
    End If
    GetFileDetails filename
End Function

Public Function HasBeenModified(filename As String, OldDate As String) As Boolean
    Dim temp As String, temp2 As Long
    temp = FileDateTime(filename)
    temp2 = DateDiff("s", temp, OldDate)
    HasBeenModified = temp2 < 0
End Function

Public Sub AddFileDetails(filename As String)
    Dim temp As String
    temp = Replace(filename, "\", "|")
    ReadID3 filename, MP3Info
    With MP3Info
        HN.creationsection "Songs\" & temp
        HN.setkeycontents "Songs\" & temp, "Date", Now
        If Len(.sArtist) > 0 Then HN.setkeycontents "Songs\" & temp, "Artist", .sArtist
        If Len(.sAlbum) > 0 Then HN.setkeycontents "Songs\" & temp, "Album", .sAlbum
        If Len(.sGenre) > 0 Then HN.setkeycontents "Songs\" & temp, "Genre", .sGenre
        If Len(.sTitle) > 0 Then HN.setkeycontents "Songs\" & temp, "Title", .sTitle
        If .sTrack > 0 Then HN.setkeycontents "Songs\" & temp, "Track", .sTrack & Empty
        If Len(.sYear) > 0 Then HN.setkeycontents "Songs\" & temp, "Year", .sYear
        If Len(.sComment) > 0 Then HN.setkeycontents "Songs\" & temp, "Comment", .sComment
        
        If Len(.sArtist) > 0 Then
            HN.creationsection "Artists\" & .sArtist
            HN.setkeycontents "Artists\" & .sArtist, temp, Empty
        End If
        
        If Len(.sAlbum) > 0 Then
            HN.creationsection "Albums\" & .sAlbum
            HN.setkeycontents "Albums\" & .sAlbum, temp, Empty
        End If
        
        If Len(.sGenre) > 0 Then
            HN.creationsection "Genres\" & .sGenre
            HN.setkeycontents "Genres\" & .sGenre, temp, Empty
        End If
    End With
End Sub
Public Sub GetFileDetails(filename As String)
    Dim temp As String
    temp = "Songs\" & Replace(filename, "\", "|")
    With MP3Info
        .sArtist = HN.getkeycontents(temp, "Artist")
        .sAlbum = HN.getkeycontents(temp, "Album")
        .sGenre = HN.getkeycontents(temp, "Genre")
        .sTitle = HN.getkeycontents(temp, "Title")
        .sTrack = HN.getkeycontents(temp, "Track")
        .sYear = HN.getkeycontents(temp, "Year")
        .sComment = HN.getkeycontents(temp, "Comment")
    End With
End Sub
Public Sub DeleteFileDetails(filename As String)
    Dim temp As String
    temp = Replace(filename, "\", "|")
    GetFileDetails filename
    With MP3Info
        HN.deletekey "Artists\" & .sArtist, temp
        HN.deletekey "Albums\" & .sAlbum, temp
        HN.deletekey "Genres\" & .sGenre, temp
        HN.deletekey "Scanned", filename
        
        If HN.keycount("Artists\" & .sArtist) = 0 Then HN.deletesection "Artists\" & .sArtist
        If HN.keycount("Albums\" & .sAlbum) = 0 Then HN.deletesection "Albums\" & .sAlbum
        If HN.keycount("Genres\" & .sGenre) = 0 Then HN.deletesection "Genres\" & .sGenre
        
        HN.deletesection "Songs\" & temp
    End With
End Sub

Public Sub ListSection(ByVal path As String, Mnumain)
    Dim tempstr() As String, temp As Long, count As Long
    path = LCase(path)
    Select Case path
        Case "songs", "artists", "albums", "genres"
            HN.enumeratesections path, tempstr
            temp = HN.sectioncount(path)
            For count = 1 To temp
                If path = "songs" Then
                    Mnumain.NewItem GetSongTitle(tempstr(count))
                Else
                    Mnumain.NewItem tempstr(count), ">"
                End If
            Next
        Case Else
            If InStr(path, "\") = InStrRev(path, "\") Then
                HN.enumeratekeys path, tempstr
                temp = HN.keycount(path)
                For count = 1 To temp
                    Mnumain.NewItem GetSongTitle(tempstr(1, count))
                Next
            End If
    End Select
End Sub
Public Function GetSongTitle(filename As String)
    Dim temp As String
    temp = Replace(filename, "\", "|")
    GetSongTitle = HN.getkeycontents("Songs\" & temp, "Title")
End Function
Public Sub ListSections(path As String, Mnumain, Optional rside As String)
    Dim tempstr() As String, temp As Long, count As Long
    HN.enumeratesections path, tempstr
    temp = HN.sectioncount(path)
    For count = 1 To temp
        Mnumain.NewItem tempstr(count), rside
    Next
End Sub
Public Sub ListKeys(path As String, Mnumain, index As Long, Optional rside As String)
    Dim tempstr() As String, temp As Long, count As Long
    HN.enumeratekeys path, tempstr
    temp = HN.keycount(path)
    For count = 1 To temp
        Mnumain.NewItem tempstr(index, count), rside
    Next
End Sub
Public Function GetFilename(path As String) As String
    Dim tempstr() As String, tempstr2() As String, temp As Long, count As Long, key As String
    tempstr2 = Split(path, "\")
    tempstr2(0) = Left(tempstr2(0), Len(tempstr2(0)) - 1)
    If UBound(tempstr2) < 2 Then Exit Function
    HN.enumeratesections "Songs", tempstr
    temp = HN.sectioncount("Songs")
    
    For count = 1 To temp
        If StrComp(tempstr2(1), HN.getkeycontents("Songs\" & tempstr(count), tempstr2(0)), vbTextCompare) = 0 Then
            If StrComp(tempstr2(2), HN.getkeycontents("Songs\" & tempstr(count), "Title"), vbTextCompare) = 0 Then
                GetFilename = Replace(tempstr(count), "|", "\")
            End If
        End If
    Next
End Function
