VERSION 5.00
Begin VB.UserControl ThumbWheel 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1155
   MaskColor       =   &H000000FF&
   MaskPicture     =   "ThumbWheel.ctx":0000
   OLEDropMode     =   1  'Manual
   Picture         =   "ThumbWheel.ctx":43F2
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   77
   ToolboxBitmap   =   "ThumbWheel.ctx":87E4
   Begin VB.Timer Timerthumb 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   360
   End
   Begin VB.Timer Timermain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   360
   End
   Begin VB.Image imgmain 
      Height          =   1800
      Index           =   1
      Left            =   2400
      Picture         =   "ThumbWheel.ctx":8AF6
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image imgmain 
      Height          =   1140
      Index           =   0
      Left            =   1200
      Picture         =   "ThumbWheel.ctx":133F8
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "ThumbWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const Pi As Double = 3.14159265358979
Private oldangle As Long, newangle As Long, pic As Long, stilldown As Boolean

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDrag(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event PodDown(Button As Integer, Shift As Integer, Angle As Long)
Public Event PodMove(Button As Integer, Shift As Integer, Angle As Long)
Public Event PodDrag(Button As Integer, Shift As Integer, Angle As Long)
Public Event PodUp(Button As Integer, Shift As Integer, Angle As Long)

Public Event PodChange(Angle As Long)
Public Event PodChangeClockWise(Angle As Long)
Public Event PodChangeCounterClockWise(Angle As Long)

Public Event ThumbClick()
Public Event ThumbStillDown()
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Public Function MouseIsDown() As Boolean
    MouseIsDown = TimerMain.Enabled
End Function
Public Property Let Size(temp As Boolean)
    pic = 1
    If temp = False Then pic = 0
    UserControl.Picture = imgmain(pic).Picture
    UserControl.MaskPicture = UserControl.Picture
    UserControl_Resize
End Property
Public Property Get Size() As Boolean
    Size = pic = 1
End Property

Public Property Let Interval(MilliSeconds As Long)
    If MilliSeconds > 0 And MilliSeconds <= 1000 Then
        TimerMain.Interval = MilliSeconds
    End If
End Property
Public Property Get Interval() As Long
    Interval = TimerMain.Interval
End Property
Public Property Let ThumbInterval(MilliSeconds As Long)
    If MilliSeconds > 0 And MilliSeconds <= 10000 Then
        Timerthumb.Interval = MilliSeconds
    End If
End Property
Public Property Get ThumbInterval() As Long
    ThumbInterval = Timerthumb.Interval
End Property

Private Function Center() As Single
    Center = UserControl.Width / 30 'divide by 15 to convert to pixels, divide by 2 to get mid point
End Function

Private Sub TimerMain_Timer()
    Dim temp As Long
    temp = newangle - oldangle
    If temp > 270 Then temp = -360 + temp 'Crossed over the 0 degree line
    If temp < -270 Then temp = 360 + temp 'Must invert the difference
    RaiseEvent PodChange(temp)
    If temp < 0 Then RaiseEvent PodChangeClockWise(Abs(temp))
    If temp > 0 Then RaiseEvent PodChangeCounterClockWise(temp)
    oldangle = newangle
End Sub

Private Sub Timerthumb_Timer()
    If Not stilldown Then
        stilldown = True
        RaiseEvent ThumbStillDown
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp As Single
    temp = Center
    temp = Distance(X, Y, temp, temp)
    If temp < IIf(Size, 20, 13) Then
        stilldown = False
        Timerthumb.Enabled = True
    Else
        TimerMain.Enabled = True
    End If
    UserControl_MouseMove Button, Shift, X, Y
    RaiseEvent MouseDown(Button, Shift, X, Y)
    RaiseEvent PodDown(Button, Shift, GetAngle(X, Y))
    oldangle = GetAngle(X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    RaiseEvent PodMove(Button, Shift, GetAngle(X, Y))
    If TimerMain.Enabled Then
        RaiseEvent MouseDrag(Button, Shift, X, Y)
        RaiseEvent PodDrag(Button, Shift, GetAngle(X, Y))
        newangle = GetAngle(X, Y)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not stilldown And Timerthumb.Enabled Then RaiseEvent ThumbClick
    TimerMain.Enabled = False
    Timerthumb.Enabled = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
    RaiseEvent PodUp(Button, Shift, GetAngle(X, Y))
End Sub

Private Function GetAngle(X As Single, Y As Single) As Long
    Dim temp As Single
    temp = Center
    GetAngle = AngleBySection(X, Y, temp, temp, RadiansToDegrees(Angle(X, Y, temp, temp)))
End Function
Private Function Angle(x1 As Single, Y1 As Single, x2 As Single, Y2 As Single) As Double
    On Error Resume Next
    Angle = Atn((Y2 - Y1) / (x1 - x2))
End Function
Private Function RadiansToDegrees(ByVal Radians As Double) As Double 'Converts Radians to Degrees.
    RadiansToDegrees = Radians * (180 / Pi)
End Function
Private Function DegreesToRadians(ByVal Degrees As Double) As Double 'Converts Degrees to Radians.
    DegreesToRadians = Degrees * (Pi / 180)
End Function
Private Function AngleBySection(x1 As Single, Y1 As Single, x2 As Single, Y2 As Single, ByVal Angle As Long) As Double
    Dim temp As Single
    temp = Center
    Angle = Abs(Angle)
    
    If x1 < x2 Then 'the point is at the left of Center
        If Y1 = Y2 Then AngleBySection = 180
        If Y1 < Y2 Then AngleBySection = 180 - Angle
        If Y1 > Y2 Then AngleBySection = 180 + Angle
    End If
    
    If x1 > x2 Then 'the point is at the right of Center
        If Y1 > Y2 Then AngleBySection = 360 - Angle
        If Y1 < Y2 Then AngleBySection = Angle
    End If
    
    If x1 = x2 Then
        If Y1 < Y2 Then AngleBySection = 90
        If Y1 > Y2 Then AngleBySection = 270
    End If
End Function

Private Sub UserControl_Resize()
    UserControl.Width = imgmain(pic).Width * 15
    UserControl.Height = UserControl.Width
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ThumbInterval = PropBag.ReadProperty("ThumbInterval", 1000)
    Interval = PropBag.ReadProperty("Interval", 100)
    Size = PropBag.ReadProperty("Size", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ThumbInterval", Timerthumb.Interval, 1000
    PropBag.WriteProperty "Interval", TimerMain.Interval, 100
    PropBag.WriteProperty "Size", pic = 1, False
End Sub

Public Function Distance(x1 As Single, Y1 As Single, x2 As Single, Y2 As Single) As Single
    On Error Resume Next
    If Y2 - Y1 = 0 Then Distance = Abs(x2 - x1): Exit Function
    If x2 - x1 = 0 Then Distance = Abs(Y2 - Y1): Exit Function
    Distance = Abs(Y2 - Y1) / Sin(Atn(Abs(Y2 - Y1) / Abs(x2 - x1)))
End Function
