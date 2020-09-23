VERSION 5.00
Begin VB.UserControl iPodMenu 
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   194
   ToolboxBitmap   =   "Menu.ctx":0000
   Begin IPod.LCD LCDmain 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3625
      Border          =   0   'False
      Begin IPod.ScrollBar scrmain 
         Height          =   1815
         Left            =   2400
         Top             =   120
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   3201
         Max             =   4
         Value           =   4
         LargeChange     =   1
      End
   End
End
Attribute VB_Name = "iPodMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ItemHeight As Long = 18, WipeSpeed As Long = 10
Private Type MenuItem
     Lside As String
     rside As String
End Type
Private MenuCount As Long, MenuList() As MenuItem, dir As Boolean, inter As Long
Private SelItem As Long, start As Long, onScreen As Long, locke As Boolean
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Function ItemCount() As Long
    ItemCount = MenuCount
End Function
Private Sub LCDmain_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub LCDmain_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub LCDmain_KeyUp(KeyCode As Integer, Shift As Integer)
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
Public Function hwnd() As Long
    hwnd = UserControl.hwnd
End Function
Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub LCDmain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Public Property Let Interval(temp As Long)
    inter = temp
End Property
Public Property Get Interval() As Long
    Interval = inter
End Property
Public Property Let Locked(temp As Boolean)
    locke = temp
    If temp = False Then DrawMenu
End Property
Public Property Get Locked() As Boolean
    Locked = locke
End Property
Public Property Let direction(temp As Boolean)
    dir = temp
End Property
Public Property Get direction() As Boolean
    direction = dir
End Property
Public Property Let BackColor(temp As OLE_COLOR)
    LCDmain.BackColor = temp
    UserControl.BackColor = temp
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = LCDmain.BackColor
End Property
Public Function GetItem(index As Long, side As Boolean) As String
    If index >= 0 And index < MenuCount Then GetItem = IIf(side, MenuList(index).Lside, MenuList(index).rside)
End Function
Public Function NewItem(text As String, Optional side As String) As Long
    NewItem = MenuCount
    With scrmain
        .Max = MenuCount
        .LargeChange = onScreen
    End With
    MenuCount = MenuCount + 1
    ReDim Preserve MenuList(MenuCount)
    SetItem MenuCount - 1, text, side
End Function
Public Sub SetItem(index As Long, text As String, Optional side As String)
    If index >= 0 And index < MenuCount Then
        With MenuList(index)
            .Lside = text
            .rside = side
        End With
    End If
    If index >= start And index <= start + onScreen - 1 Then DrawMenu
End Sub
Public Sub ClearItems(Optional DoWipe As Boolean = False, Optional DoUnWipe As Boolean)
    MenuCount = 0
    ReDim MenuList(0)
    With scrmain
        .value = 0
        .Max = 0
        .LargeChange = 0
    End With
    SelItem = 0
    start = 0
    If DoWipe Then Wipe
    LCDmain.ClearText
    If DoUnWipe Then UnWipe
    DoEvents
End Sub
Public Sub Wipe()
    If inter = 0 Then inter = WipeSpeed
    Dim temp As Long, wid As Long
    wid = UserControl.Width
    For temp = 0 To wid / 15 Step inter
        If dir Then 'left
            LCDmain.Left = LCDmain.Left - inter
            scrmain.Left = scrmain.Left - inter
        Else 'right
            LCDmain.Left = LCDmain.Left + inter
            scrmain.Left = scrmain.Left + inter
        End If
        DoEvents
    Next
End Sub
Public Sub Pacman()
    Dim wid As Long
    wid = UserControl.Width / 15
    If dir = True Then
        LCDmain.Left = wid
        scrmain.Left = wid + LCDmain.Width
    Else
        LCDmain.Left = 0 - LCDmain.Width - scrmain.Width
        scrmain.Left = 0 - scrmain.Width
    End If
    dir = Not dir
    UnWipe
    dir = Not dir
End Sub

Public Sub UnWipe()
    Dim temp As Long, wid As Long
    wid = UserControl.Width
    If inter = 0 Then inter = WipeSpeed
    For temp = 0 To wid / 15 Step inter
        If LCDmain.Left < 0 Then
            LCDmain.Left = LCDmain.Left + inter
            scrmain.Left = scrmain.Left + inter
        Else 'right
            LCDmain.Left = LCDmain.Left - inter
            scrmain.Left = scrmain.Left - inter
        End If
        DoEvents
    Next
    'MsgBox LCDmain.Left
    LCDmain.Left = 0
    scrmain.Move wid - scrmain.Width
End Sub
Public Property Let SelectedItem(index As Long)
    If index < 0 Then index = 0
    If index >= MenuCount Then index = MenuCount - 1
    If SelItem = index Then Exit Property
    SelItem = index
    scrmain.value = index
    If index > start + onScreen - 1 Then
        Do Until index <= start + onScreen - 1
            start = start + 1
        Loop
    End If
    If index < start Then start = index
    DrawMenu
End Property
Public Property Get SelectedItem() As Long
    SelectedItem = SelItem
End Property

Private Sub UserControl_Resize()
    Dim hit As Long, wid As Long
    UserControl.Height = (((UserControl.Height / 15) \ ItemHeight) * ItemHeight) * 15
    
    hit = UserControl.Height
    wid = UserControl.Width
    
    LCDmain.Move 0, 0, wid, hit
    scrmain.Move wid - scrmain.Width, 0, scrmain.Width, hit
    
    onScreen = (hit / 15) \ ItemHeight
    With scrmain
        .Max = MenuCount
        .LargeChange = onScreen
    End With
    DrawMenu
End Sub
Public Sub DrawMenu()
    If locke Then Exit Sub
    Dim temp As Long, Y As Long, wid As Long, templ As Long, tempr As Long, hit As Long
    If MenuCount = 0 Then Exit Sub
    Const WhiteSpace As Long = 4
    wid = (UserControl.Width - scrmain.Width) / 15
    LCDmain.ClearText
    For temp = start To start + onScreen - 1
        If temp >= 0 And temp < MenuCount Then
            hit = StringHeight(MenuList(temp).Lside & MenuList(temp).rside)
            Y = (temp - start) * ItemHeight
            tempr = StringWidth(MenuList(temp).rside) + WhiteSpace
            If temp = SelItem Then
                templ = StringWidth(MenuList(temp).Lside) + WhiteSpace
                LCDmain.DrawSquare 0, Y, wid, 4, vbBlack, True 'Top Bar
                LCDmain.DrawSquare 0, Y + ItemHeight - 2, wid, 2, vbBlack, True 'Bottom bar
                LCDmain.DrawSquare 0, Y + WhiteSpace, 4, 12, vbBlack, True 'Left side middle bar
                LCDmain.DrawSquare templ, Y, wid - templ - tempr, ItemHeight, vbBlack, True    'Middle middle bar
                LCDmain.DrawSquare wid - WhiteSpace, Y + WhiteSpace, WhiteSpace, ItemHeight - WhiteSpace, vbBlack, True
            End If
            LCDmain.PrintText Truncate(MenuList(temp).Lside, wid - WhiteSpace - WhiteSpace), WhiteSpace, Y + 4, temp = SelItem
            LCDmain.PrintText MenuList(temp).rside, wid - tempr, Y + 4, temp = SelItem
        End If
    Next
    DoEvents
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BackColor = PropBag.ReadProperty("BackColor", &HC8DDC1)
    direction = PropBag.ReadProperty("Direction", False)
    Interval = PropBag.ReadProperty("Interval", WipeSpeed)
    Locked = PropBag.ReadProperty("Locked", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", LCDmain.BackColor, &HC8DDC1
    PropBag.WriteProperty "Direction", dir, False
    PropBag.WriteProperty "Interval", inter, WipeSpeed
    PropBag.WriteProperty "Locked", locke, False
End Sub
