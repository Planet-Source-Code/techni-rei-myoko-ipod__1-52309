Attribute VB_Name = "iPodMenus"
Option Explicit
Public Enum Ipod_Mode
    ipod_menu = 0
    ipod_nowplaying = 1
    ipod_block = 2
    ipod_parachute = 3
End Enum
'Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "User32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
    'Used to set window to always be on top or not
    Public Const HWND_NOTOPMOST = -2
    Public Const HWND_TOPMOST = -1
    
Public CurrDir As String, OnTop As Boolean
Public DrvBox As DriveListBox, DirBox As DirListBox, FilBox As FileListBox, lC As Object, Bat As Object, Bar As Object, tim As Object
Public PlayList() As String, PlayCount As Long, CurrItem As Long, TitleCaption As String, MenuMode As Long
Public Function chkpath(path As String, filename As String) As String
    chkpath = path & IIf(InStrRev(path, "\") = Len(path), Empty, "\") & filename
End Function
Public Sub setAlwaysOnTop(hwnd As Long, Optional OnTop As Boolean = True)
On Error Resume Next
If OnTop = False Then Call SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
If OnTop = True Then Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
End Sub
Public Sub Seekto(direction As Long)
    Dim isdone As Boolean
    If MediaIsPlaying And direction = -1 Then
        If MediaCurrentPosition > 4 Then
            MediaSeekto 0
            isdone = True
        End If
    End If
    If PlayCount > 0 And isdone = False Then
        CurrItem = CurrItem + direction
        If CurrItem < 0 Then CurrItem = PlayCount - 1
        If CurrItem >= PlayCount Then CurrItem = 0
        PlayItem CurrItem
    End If
End Sub
Public Function isadir(filename As String) As Boolean
On Error Resume Next
If filename <> Empty Then isadir = (GetAttr(filename) And vbDirectory) = vbDirectory
End Function
Public Function fileexists(filename As String) As Boolean
    On Error Resume Next
    Dim temp As Long
    temp = FileLen(filename)
    fileexists = temp > 0
End Function
Public Sub AddPlayItem(filename As String)
    PlayCount = PlayCount + 1
    ReDim Preserve PlayList(PlayCount)
    PlayList(PlayCount - 1) = filename
End Sub
Public Sub AddPlayList(filename As String, Optional index As Long = -1)
    Dim tempstr() As String, tempcount As Long, temp As Long
    LoadPlayList filename, tempstr, tempcount
    For temp = 0 To tempcount - 1
        AddPlayItem tempstr(temp)
    Next
    If index > -1 Then CurrItem = PlayCount + index
End Sub
Public Sub ErasePlayList(index As Long, range As Long)
    Dim temp As Long
    If CurrItem >= index And CurrItem < index + range Then
        MediaStop
        CurrItem = index - 1
    End If
    For temp = index To PlayCount - range
        PlayList(temp) = PlayList(temp) + range
    Next
    PlayCount = PlayCount - range
    ReDim Preserve PlayList(PlayCount)
End Sub
Public Sub PlayItem(ByVal index As Long)
    If PlayCount = 0 Then Exit Sub
    If index < 0 Then
        Do Until index >= 0
            index = index + PlayCount
        Loop
    End If
    If index >= PlayCount Then index = index Mod PlayCount
    If CurrItem = index Or MediaIsPlaying Then MediaSeekto 0
    CurrItem = index
    MediaLoad PlayList(index)
    ScanFile PlayList(index)
    MediaPlay
End Sub
Public Sub ClearPlayList()
    PlayCount = 0
    ReDim PlayList(0)
    CurrItem = 0
End Sub

Public Sub TitleBar(LCDmain, text As String, Optional Playstate As String)
    Dim temp As Long
    Static OldText As String
    If Len(text) = 0 Then text = OldText Else OldText = text
    temp = StringWidth(text)
    LCDmain.PrintText text, (LCDmain.Width - temp) / 2, 5
    LCDmain.PrintText Playstate, 13, 5
    LCDmain.DrawLine 2, 20, LCDmain.Width - 7, 1
    TitleCaption = text
End Sub
Public Sub ShowPlayList(mnu, filename As String)
Dim tempstr() As String, tempcount As Long, temp As Long
If filename = "on the go" Then
    For temp = 0 To PlayCount - 1
        mnu.NewItem afterslash(PlayList(temp))
    Next
Else
    LoadPlayList filename, tempstr, tempcount
    For temp = 0 To tempcount - 1
        mnu.NewItem afterslash(tempstr(temp))
    Next
End If
End Sub
Public Function afterslash(text As String) As String
    If InStr(text, "\") = 0 Then
        afterslash = text
    Else
        afterslash = Right(text, Len(text) - InStrRev(text, "\"))
    End If
End Function
Public Sub MainMenu(lcdm, mnu, tim, Optional ByVal location As String, Optional index As Long = -1)
'On Error Resume Next
    With mnu
        .Locked = True
        .top = 390
        .Height = 1620
        MenuMode = ipod_menu
        lcdm.ClearText
        .ClearItems True
        location = LCase(Trim(location))
        If Right(location, 1) = "\" Then location = Left(location, Len(location) - 1)
        If Left(location, 12) = "now playing\" Then location = Right(location, Len(location) - 12)
        Select Case location
            Case Empty
                TitleBar lcdm, "MyPod", PlayPause
                .NewItem "Playlists", ">"
                .NewItem "Browse", ">"
                .NewItem "Last Played", ">"
                .NewItem "Settings", ">"
                .NewItem "Extra", ">"
            
            Case "playlists"
                TitleBar lcdm, "Playlists", PlayPause
                ListKeys "Playlists", mnu, 2, ">"
                .NewItem "On the Go", ">"
            
            Case "browse"
                TitleBar lcdm, "Browse", PlayPause
                .NewItem "Artists", ">"
                .NewItem "Albums", ">"
                .NewItem "Songs", ">"
                .NewItem "Genres", ">"
                .NewItem "System", ">"
                
            Case "extra"
                TitleBar lcdm, "Extra", PlayPause
                .NewItem "Clock", ">"
                .NewItem "Contacts", ">"
                .NewItem "Calender", ">"
                .NewItem "Notes", ">"
                .NewItem "Games", ">"
                
                Case "extra\clock"
                    .Height = 810
                    .top = 1200
                    TitleBar lcdm, Format(Date, "MMM d YYYY"), PlayPause
                    .NewItem "Alarm Clock", ">"
                    .NewItem "Sleep and Timer", ">"
                    .NewItem "Date & Time"
                    tim.Enabled = True
            
                    Case "extra\clock\alarm clock"
                        TitleBar lcdm, "Alarm Clock", PlayPause
                        .NewItem "Alarm", "On"
                        .NewItem "Time", ">"
                        .NewItem "Sound"
            
                Case "extra\calender"
                    TitleBar lcdm, "Calender", PlayPause
                    .NewItem "All"
                    .NewItem "To do"
                    .NewItem "Alarms"
            
                Case "extra\games"
                    TitleBar lcdm, "Games", PlayPause
                    .NewItem "Brick"
                    .NewItem "Parachute"
                    '.NewItem "Solitaire"
                
                Case "extra\contacts"
                    TitleBar lcdm, "Contacts", PlayPause
                    ListSections "Contacts", mnu, ">"
                
            Case "settings"
                TitleBar lcdm, "Settings", PlayPause
                .NewItem "About", ">"
                .NewItem "Main Menu", ">"
                .NewItem "Shuffle", "Off"
                .NewItem "Repeat", "On"
                .NewItem "Always on Top", IIf(OnTop, "On", "Off")
                .NewItem "Contrast", "Middle"
                .NewItem "Clicker", "On"
                .NewItem "Sleep Timer", "Off"
                .NewItem "Startup Volume", "Middle"
                .NewItem "Alarms", "On" 'Silent (visual), off
                
                Case "settings\about"
                    TitleBar lcdm, "About", PlayPause
                    .NewItem "Songs", "0"
                    .NewItem "Capacity", "20 GB"
                    .NewItem "Available", "104 MB"
                    .NewItem "Version", "1.3"
                    .NewItem "S/N", "U" & Round(Rnd * 100000) & "TRM"
                    .NewItem "Format", "Windows"

            Case Else
                If isinSection("Playlists", location) Then
                    If location <> "playlists\on the go" Then
                        CurrPlayList = GetPlaylist(index + 1)
                        TitleBar lcdm, afterslash(CurrPlayList), PlayPause
                        ShowPlayList mnu, CurrPlayList
                    Else
                        TitleBar lcdm, "On the Go", PlayPause
                        ShowPlayList mnu, "on the go"
                    End If
                End If
                If isinSection("Browse", location) Then
                    If isinSection("Browse\System", location) Then
                        Dim temp As String, count As Long
                        location = Replace(location, "<dir>", Empty, , , vbTextCompare)
                        If StrComp(location, "Browse\System", vbTextCompare) <> 0 Then temp = Right(location, Len(location) - 14)
                        If Len(temp) = 0 Then
                            TitleBar lcdm, "System", PlayPause
                            For count = 0 To DrvBox.ListCount - 1
                                .NewItem "<dir>" & Left(DrvBox.List(count), 2), ">"
                            Next
                        Else
                            TitleBar lcdm, afterslash(temp), PlayPause
                            If Len(temp) = 2 Then temp = temp & "\"
                            DirBox.path = temp
                            FilBox.path = temp
                            DumpContents mnu, DirBox, "<dir>", ">"
                            FilBox.Pattern = PlayListFiles
                            DumpContents mnu, FilBox, Empty
                            FilBox.Pattern = AudioFiles
                            DumpContents mnu, FilBox, Empty
                        End If
                    Else
                        location = Right(location, Len(location) - 7)
                        TitleBar lcdm, afterslash(sCase(location)), PlayPause
                        ListSection location, mnu
                    End If
                End If
        End Select
        .Locked = False
        .Pacman
    End With
End Sub
Public Function sCase(text As String) As String
    sCase = UCase(Left(text, 1)) & LCase(Right(text, Len(text) - 1))
End Function
Public Sub DumpContents(Mnumain, Obj, Optional Icon As String, Optional Rightside As String)
    Dim temp As Long, tempstr As String
    For temp = 0 To Obj.ListCount - 1
        tempstr = Obj.List(temp)
        If InStr(tempstr, "\") > 0 Then tempstr = Right(tempstr, Len(tempstr) - InStrRev(tempstr, "\"))
        Mnumain.NewItem Icon & tempstr, Rightside
    Next
End Sub
Public Function isinSection(current As String, section As String) As Boolean
    isinSection = StrComp(current, Left(section, Len(current)), vbTextCompare) = 0
End Function
Public Function GetMenu(ByVal current As String, Relative As String) As String
    If Right(current, 1) = "\" Then current = Left(current, Len(current) - 1)
    If Left(Relative, 1) = "\" Then Relative = Right(Relative, Len(Relative) - 1)
    If Relative = "..\" Then
        If Len(current) > 0 Then
            If InStrRev(current, "\") > 0 Then
                GetMenu = Left(current, InStrRev(current, "\") - 1)
            End If
        End If
    Else
        GetMenu = current & "\" & Relative
        If Len(current) = 0 Then GetMenu = Relative
        If Len(Relative) = 0 Then GetMenu = current
    End If
End Function

Public Sub Execute(frm As Form, mnu, ByVal current As String, filename As String, Optional index As Long)
With mnu
    If isinSection("Extra\games", current) Then
        Select Case LCase(filename)
            Case "brick"
                MenuMode = ipod_block
                InitializeDefaults lC
            Case "parachute"
                MenuMode = ipod_parachute
                tim.Interval = 100
                InitParachuteVars
            Case Else
                MsgBox filename & " is not available as of yet"
        End Select
        mnu.Visible = False
        Bar.Visible = False
        Bat.Visible = False
        tim.Enabled = True
    End If
    If isinSection("Browse", current) Then
        If isinSection("Browse\System", current) Then
            current = Right(current, Len(current) - 14)
            current = GetMenu(current, filename)
        Else
            current = current & "\" & filename
            current = Right(current, Len(current) - 7)
            current = GetFilename(current)
        End If
        If Len(current) > 0 Then
            current = Replace(current, "<dir>", Empty, , , vbTextCompare)
            If isaPlaylist(current) Then
                CurrItem = PlayCount
                AddPlayList current
                NewPlaylist current
                PlayItem CurrItem
                MenuMode = ipod_nowplaying
            Else
                AddPlayItem current
                PlayItem PlayCount - 1
                MenuMode = ipod_nowplaying
            End If
        End If
    End If
    If isinSection("Settings", current) Then
        If StrComp(.GetItem(.SelectedItem, True), "Always on Top", vbTextCompare) = 0 Then
               OnTop = Not OnTop
               .SetItem .SelectedItem, "Always on Top", IIf(OnTop, "On", "Off")
               DoOnTop frm
        End If
    End If
    If isinSection("Playlists\On the Go", current) Then
        PlayItem index
        MenuMode = ipod_nowplaying
    Else
        If isinSection("Playlists", current) Then
            AddPlayList CurrPlayList, index
            PlayItem CurrItem
            MenuMode = ipod_nowplaying
        End If
    End If
End With
End Sub
Public Sub DoOnTop(frm As Form)
    setAlwaysOnTop frm.hwnd, OnTop
End Sub
Public Function isaPlaylist(filename As String) As Boolean
    isaPlaylist = islike(PlayListFiles, filename)
End Function
Public Function islike(filter As String, ByVal expression As String) As Boolean
Dim tempstr() As String, count As Long
tempstr = Split(LCase(filter), ";")
expression = LCase(expression)
islike = False
For count = 0 To UBound(tempstr)
    If expression Like tempstr(count) Then islike = True: Exit For
Next
End Function
Public Function Truncate(ByVal text As String, Width As Long) As String
    If StringWidth(text) <= Width Then
        Truncate = text
    Else
        Do Until StringWidth(text & "...") <= Width
            text = Left(text, Len(text) - 1)
        Loop
        Truncate = text & "..."
    End If
End Function

Public Function PlayPause() As String
    If MediaIsPlaying Then PlayPause = "<play>"
    If MediaIsPaused Then PlayPause = "<pause>"
End Function
Public Function findXY(X As Single, Y As Single, Distance As Single, Angle As Double, Optional isx As Boolean = True) As Single
    If isx = True Then findXY = X + Sin(Angle) * Distance Else findXY = Y + Cos(Angle) * Distance
End Function
Public Function rad2deg(Radians As Double) As Double
    rad2deg = Radians * 180
End Function
