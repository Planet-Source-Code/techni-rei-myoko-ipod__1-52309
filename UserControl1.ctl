VERSION 5.00
Begin VB.UserControl LCD 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   ControlContainer=   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   136
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   169
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C8DDC1&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   360
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.PictureBox Picfont 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   0
      Picture         =   "UserControl1.ctx":0312
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.Shape shpmain 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   240
      Top             =   240
      Width           =   2055
   End
   Begin VB.Line Linblack 
      Index           =   3
      X1              =   160
      X2              =   160
      Y1              =   16
      Y2              =   120
   End
   Begin VB.Line Linblack 
      Index           =   2
      X1              =   8
      X2              =   8
      Y1              =   16
      Y2              =   120
   End
   Begin VB.Line Linblack 
      Index           =   1
      X1              =   16
      X2              =   152
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line Linblack 
      Index           =   0
      X1              =   16
      X2              =   152
      Y1              =   8
      Y2              =   8
   End
End
Attribute VB_Name = "LCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const IpodBack As Long = 13163969
Private hasborder As Boolean
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Sub picmain_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picmain_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picmain_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub picmain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Public Function hwnd() As Long
    hwnd = UserControl.hwnd
End Function
Public Property Let BackColor(temp As OLE_COLOR)
    picmain.BackColor = temp
    UserControl.BackColor = temp
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = picmain.BackColor
End Property
    
Public Property Let Border(temp As Boolean)
    hasborder = temp
    UserControl_Resize
End Property
Public Property Get Border() As Boolean
    Border = hasborder
End Property

Private Sub UserControl_Resize()
    Dim Width As Long, Height As Long
    Width = UserControl.Width / 15
    Height = UserControl.Height / 15
    
    MoveLine Linblack(0), 1, 0, Width - 1, 1
    MoveLine Linblack(1), 1, Height - 1, Width - 1, 1
    MoveLine Linblack(2), 0, 1, 1, Height - 1
    MoveLine Linblack(3), Width - 1, 1, 1, Height - 1
    
    shpmain.Move 1, 1, Width - 2, Height - 2
    If hasborder Then
        picmain.Move 2, 2, Width - 4, Height - 4
    Else
        picmain.Move 0, 0, Width, Height
    End If
End Sub
Public Function hdc() As Long
    hdc = picmain.hdc
End Function
Private Sub MoveLine(lin As line, Left As Long, top As Long, Width As Long, Height As Long)
    With lin
        .x1 = Left
        .Y1 = top
        
        .x2 = Left + Width - 1
        .Y2 = top + Height - 1
    End With
End Sub

Public Sub PrintText(text As String, X As Long, Y As Long, Optional hi As Boolean)
    iPrint text, Picfont.hdc, picmain.hdc, X, Y, hi
    picmain.Refresh
End Sub
Public Sub ClearText()
    picmain.Cls
End Sub
Public Sub DrawSquare(X As Long, Y As Long, Width As Long, Height As Long, Optional color As Long = vbBlack, Optional Filled As Boolean)
    If Filled Then
        picmain.FillStyle = vbSolid
        picmain.FillColor = color
    End If
    picmain.Line (X, Y)-(X + Width - 1, Y + Height - 1), color, B
    picmain.FillStyle = vbTransparent
    picmain.Refresh
End Sub
Public Sub DrawLine(X As Long, Y As Long, Width As Long, Height As Long, Optional color As Long = vbBlack)
    picmain.Line (X, Y)-(X + Width - 1, Y + Height - 1), color
    picmain.Refresh
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Border = PropBag.ReadProperty("Border", True)
    BackColor = PropBag.ReadProperty("BackColor", IpodBack)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Border", hasborder, True
    PropBag.WriteProperty "BackColor", picmain.BackColor, IpodBack
End Sub
