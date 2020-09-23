VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "MyPod"
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   Icon            =   "frmmain1.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   Picture         =   "frmmain1.frx":0E42
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   247
   ShowInTaskbar   =   0   'False
   Begin IPod.Hini Hinmain 
      Left            =   3120
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin IPod.iPodMenu Mnumain 
      Height          =   1350
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2381
      Direction       =   -1  'True
      Interval        =   0
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
   End
   Begin IPod.ThumbWheel ThWmain 
      Height          =   1800
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   3175
      Size            =   -1  'True
   End
   Begin IPod.LCD LCDmain 
      Height          =   2055
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3625
      Begin IPod.BatteryLevel BatMain 
         Height          =   150
         Left            =   2160
         TabIndex        =   5
         Top             =   100
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   265
      End
      Begin IPod.StatusBar barmain 
         Height          =   135
         Left            =   120
         Top             =   1620
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   238
         Max             =   360
         Value           =   0
      End
   End
   Begin VB.FileListBox Filmain 
      Height          =   285
      Left            =   840
      Pattern         =   "*.wav;*.mp3;*.wma;*.wax;*.mid;*.midi;*.rmi;*.au;*.snd;*.aif;*.aifc;*.aiff"
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.DirListBox Dirmain 
      Height          =   315
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.DriveListBox Drvmain 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   2520
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      Image           =   1
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   2
      Left            =   1920
      TabIndex        =   10
      Top             =   2520
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      Image           =   2
      Interval        =   2000
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   3
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      Image           =   3
   End
   Begin VB.Timer TimerMain 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   840
      Top             =   1320
   End
   Begin VB.PictureBox PicParachute 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   1440
      Picture         =   "frmmain1.frx":2262
      ScaleHeight     =   720
      ScaleWidth      =   1050
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1050
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents systray As clsSysTray
Attribute systray.VB_VarHelpID = -1
Private doseek As Boolean
Const iPod_Green As Long = &HC8DDC1
Const MobilePhile_Blue As Long = 13514752

Private Sub BatMain_Click()
Form_Unload 0
End Sub

Private Sub btnmain_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub btnmain_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Public Sub btnmain_MouseClick(index As Integer)
    Select Case index
        Case 0
            If MenuMode = ipod_block Or MenuMode = ipod_parachute Then
                TimerMain.Interval = TimerMain.Interval + 50
                If TimerMain.Interval > 3000 Then TimerMain.Interval = 3000
            Else
                PlayItem (CurrItem - 1)
            End If
        Case 1
            Select Case MenuMode
                Case ipod_menu
                    If Len(CurrDir) = 0 Then
                        MenuMode = ipod_nowplaying
                        TimerMain.Enabled = True
                        Mnumain.Visible = False
                    Else
                        Dim temp As String
                        temp = CurrDir
                        CurrDir = GetMenu(CurrDir, "..\")
                        If StrComp(CurrDir, temp, vbTextCompare) <> 0 Then
                            Mnumain.direction = False
                            MainMenu LCDmain, Mnumain, TimerMain, CurrDir
                            Mnumain.direction = True
                        End If
                    End If
                Case Else
                        LCDmain.ClearText
                        If InStr(CurrDir, "\") > 0 Then
                            TitleBar LCDmain, Right(CurrDir, Len(CurrDir) - InStrRev(CurrDir, "\")), PlayPause
                        Else
                            TitleBar LCDmain, CurrDir, PlayPause
                        End If
                        barmain.Visible = True
                        BatMain.Visible = True
                        Mnumain.Visible = True
                        TimerMain.Enabled = False
                        MenuMode = ipod_menu
            End Select
        Case 2
            If MenuMode = ipod_block Or MenuMode = ipod_parachute Then
                TimerMain.Enabled = Not TimerMain.Enabled
            Else
                MediaPlay
                TimerMain.Enabled = True
                Mnumain.Visible = False
                MenuMode = ipod_nowplaying
                If MediaIsPaused Then
                    TimerMain_Timer
                    TimerMain.Enabled = False
                End If
            End If
            
        Case 3
            If MenuMode = ipod_block Or MenuMode = ipod_parachute Then
                If TimerMain.Interval > 60 Then TimerMain.Interval = TimerMain.Interval - 50
            Else
                PlayItem CurrItem + 1
            End If
    End Select
End Sub

Private Sub btnmain_MouseStillDown(index As Integer, X As Single, Y As Single)
If index = 2 Then BatMain_Click
End Sub

Private Sub btnmain_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
DoEvents
End Sub
Public Sub MoveMNUmain()
    'MNUmain is not a child of LCDmain by default because:
    'Randomly, VB will choose to think LCDmain is NOT a container
    'And will fail to load the entire form, cause it doesnt want
    'to make mnumain a child of it
    
    'SetParent Mnumain.hWnd, LCDmain.hWnd 'This method prevents the creation of autoredraw images
    'Set Mnumain.Parent = LCDmain 'This method fails to work
    
    Set Mnumain.Container = LCDmain
    Mnumain.Move 30, 390, 2460, 1620
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 17: btnmain_MouseClick 0
    Case 8, 46: btnmain_MouseClick 1
    Case 80: btnmain_MouseClick 2
    Case 96: btnmain_MouseClick 3
    Case 13, 32: THWmain_ThumbClick
    Case 37, 38: THWmain_PodChangeCounterClockWise 15
    Case 39, 40: THWmain_PodChangeClockWise 15
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37, 38, 39, 40: ThWmain_PodUp 0, Shift, 0
End Select
End Sub

Private Sub Form_Load()
    Dim temp As String
    If Me.Picture <> 0 Then Call SetAutoRgn(Me)
    MoveMNUmain
    Set HN = Hinmain
    Set tim = TimerMain
    Set lC = LCDmain
    Set Bar = barmain
    Set Bat = BatMain
    Set DrvBox = Me.Drvmain
    Set DirBox = Me.Dirmain
    Set FilBox = Me.Filmain

    Set systray = New clsSysTray
    Set systray.SourceWindow = Me
    systray.Icon = Me.Icon
    systray.ToolTip = Me.Caption
    systray.IconInSysTray
    
    MediaContainersHwnd Me.hwnd
    Alias = "MyPod"
    
    LoadSettings
    
    If Len(command) = 0 Then
        MainMenu LCDmain, Mnumain, TimerMain
    Else
        temp = command
        If Left(temp, 1) = """" Then temp = Right(temp, Len(temp) - 1)
        If Right(temp, 1) = """" Then temp = Left(temp, Len(temp) - 1)
        AddPlayItem temp
        PlayItem PlayCount - 1
        MenuMode = ipod_nowplaying
        NowPlaying
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dragform Me.hwnd
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'Effect 1 list 7 filenames
Dim temp As Integer, temp2 As Long
temp = PlayCount
If Data.Files.count > 0 Then
For temp = 1 To Data.Files.count
    If (GetAttr(Data.Files(temp)) And vbDirectory) <> vbDirectory Then
        AddPlayItem Data.Files.Item(temp)
    Else
        AddFolder Data.Files.Item(temp)
    End If
    If temp = 1 Then
        PlayItem temp2
        MenuMode = ipod_nowplaying
        NowPlaying
    End If
Next
End If
End Sub
Public Sub AddFolder(path As String)
    Dim temp As Long
    Filmain.path = path
    For temp = 0 To Filmain.ListCount - 1
        AddPlayItem chkpath(path, Filmain.List(temp))
    Next
End Sub

Private Sub Form_Resize()
'    Me.Visible = Me.WindowState <> vbMinimized
'    App.TaskVisible = Me.Visible
End Sub

Private Sub Form_Unload(Cancel As Integer)
    systray.RemoveFromSysTray
    Set systray = Nothing
    MediaClose
    SaveSettings
    End
End Sub

Private Sub LCDmain_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub LCDmain_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub LCDmain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Mnumain_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub Mnumain_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub mnumain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub


Private Sub ThWmain_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub ThWmain_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub ThWmain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub THWmain_PodChangeClockWise(Angle As Long)
    Select Case MenuMode
        Case ipod_menu: Mnumain.SelectedItem = Mnumain.SelectedItem + 1
        Case ipod_nowplaying: barmain.value = barmain.value + Angle: doseek = True
        Case ipod_block: MovePaddle Angle / AngleFactor, 1
        Case ipod_parachute: MoveTurret -Angle
    End Select
End Sub

Private Sub THWmain_PodChangeCounterClockWise(Angle As Long)
    Select Case MenuMode
        Case ipod_menu: Mnumain.SelectedItem = Mnumain.SelectedItem - 1
        Case ipod_nowplaying: barmain.value = barmain.value - Angle: doseek = True
        Case ipod_block: MovePaddle Angle / AngleFactor, -1
        Case ipod_parachute: MoveTurret Angle
    End Select
End Sub

Private Sub ThWmain_PodUp(Button As Integer, Shift As Integer, Angle As Long)
    If MenuMode = ipod_nowplaying And doseek Then MediaSeekto barmain.value
    doseek = False
End Sub

Private Sub THWmain_ThumbClick()
    Dim temp As String, go As Boolean
    Select Case MenuMode
        Case ipod_menu
            temp = Mnumain.GetItem(Mnumain.SelectedItem, True)
            go = Mnumain.GetItem(Mnumain.SelectedItem, False) = ">"
            If go Then
                CurrDir = GetMenu(CurrDir, temp)
                MainMenu LCDmain, Mnumain, TimerMain, CurrDir, Mnumain.SelectedItem
            Else
                Execute Me, Mnumain, CurrDir, temp, Mnumain.SelectedItem
                NowPlaying
            End If
    
        Case ipod_block
            TimerMain.Interval = 10
            StartGame
            
        Case ipod_parachute
            CreateBullet
    
    End Select
    
    If MenuMode <> ipod_block And MenuMode <> ipod_parachute Then TimerMain.Interval = 800
End Sub
Public Sub NowPlaying()
        If MenuMode = ipod_nowplaying Then
            Mnumain.Visible = False
            TimerMain.Enabled = True
        End If
End Sub
Private Sub TimerMain_Timer()
Dim temp As String
If isReady And MenuMode = ipod_nowplaying Then
    LCDmain.ClearText
    TitleBar LCDmain, "Now Playing", PlayPause
    
    If Not ThWmain.MouseIsDown Then
        barmain.Max = MediaDuration
        barmain.value = MediaCurrentPosition
    End If
    
    temp = sec2time(MediaCurrentPosition)
    LCDmain.PrintText temp, 8, 120
    
    temp = "-" & sec2time(MediaTimeRemaining)
    LCDmain.PrintText temp, 160 - StringWidth(temp), 120
    
    temp = CurrItem + 1 & " of " & PlayCount
    If PlayCount = 0 Then
        temp = "0 of 0"
        If MediaIsPlaying Then temp = "Ghost file"
    End If
    LCDmain.PrintText temp, 4, 25
    
    temp = Truncate(MP3Info.sTitle, 165)
    If Len(temp) = 0 Then temp = Truncate(Right(MP3Info.sFilename, Len(MP3Info.sFilename) - InStrRev(MP3Info.sFilename, "\")), 165)
    LCDmain.PrintText temp, (169 - StringWidth(temp)) / 2, 46
    
    temp = Truncate(MP3Info.sArtist, 165)
    LCDmain.PrintText temp, (169 - StringWidth(temp)) / 2, 64
    
    temp = Truncate(MP3Info.sAlbum, 165)
    LCDmain.PrintText temp, (169 - StringWidth(temp)) / 2, 82
    
    If MediaTimeRemaining = 0 Then
        PlayItem CurrItem + 1
    End If
End If
If MenuMode = ipod_menu And StrComp(CurrDir, "Extra\Clock", vbTextCompare) = 0 Then
    LCDmain.ClearText
    temp = Format(Date, "MMM d YYYY")
    TitleBar LCDmain, temp, PlayPause
    temp = GetTime
    LCDmain.PrintText temp, (169 - StringWidth(temp)) / 2, 37
    LCDmain.DrawLine 2, 78, LCDmain.Width - 7, 1
End If
If MenuMode = ipod_block Then MovePuck
If MenuMode = ipod_parachute Then
    MoveParachuters
    MoveBullets
    DrawParachuteScreen Me.PicParachute.hdc, LCDmain
End If
End Sub

Public Sub LoadSettings()
    Hinmain.loadfile chkpath(App.path, "iPod.hini")

    OnTop = GetSetting("MyPod", "Main", "AlwaysOnTop", False)
    Me.Left = GetSetting("MyPod", "Main", "Left", 0)
    Me.top = GetSetting("MyPod", "Main", "Top", 0)
    
    frmremote.Left = GetSetting("MyPod", "Remote", "Left", Screen.Width - frmremote.Width)
    frmremote.top = GetSetting("MyPod", "Remote", "Top", Screen.Height - frmremote.Height)
    
    DoOnTop Me
End Sub
Public Sub SaveSettings()
    Hinmain.savefile chkpath(App.path, "iPod.hini")

    SaveSetting "MyPod", "Main", "AlwaysOnTop", OnTop
    SaveSetting "MyPod", "Main", "Left", Me.Left
    SaveSetting "MyPod", "Main", "Top", Me.top
    
    SaveSetting "MyPod", "Remote", "Left", frmremote.Left
    SaveSetting "MyPod", "Remote", "Top", frmremote.top
End Sub

Private Sub systray_LButtonUp()
On Error Resume Next
    App.TaskVisible = Not App.TaskVisible
    Me.Visible = App.TaskVisible
End Sub

Private Sub systray_RButtonUp()
    frmremote.Visible = Not frmremote.Visible
End Sub
Private Sub systray_LButtonDblClk()
    systray.IconInSysTray
End Sub

Private Sub systray_RButtonDblClk()
    systray.IconInSysTray
End Sub
