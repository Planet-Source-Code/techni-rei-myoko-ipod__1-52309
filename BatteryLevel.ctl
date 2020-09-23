VERSION 5.00
Begin VB.UserControl BatteryLevel 
   BackStyle       =   0  'Transparent
   ClientHeight    =   150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   MaskColor       =   &H00FFFFFF&
   MaskPicture     =   "BatteryLevel.ctx":0000
   Picture         =   "BatteryLevel.ctx":005E
   ScaleHeight     =   10
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   20
   ToolboxBitmap   =   "BatteryLevel.ctx":02F8
   Begin VB.Shape Shpmain 
      BackColor       =   &H80000008&
      BackStyle       =   1  'Opaque
      Height          =   90
      Left            =   30
      Top             =   30
      Width           =   225
   End
End
Attribute VB_Name = "BatteryLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Change()
Public Event Click()

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

Public Property Get Percent() As Long
    Percent = shpmain.Width / 15 * 100
End Property
Public Property Let Percent(temp As Long)
    If temp >= 0 And temp <= 100 Then Power = temp * 0.15
End Property
Public Property Let Power(temp As Long)
    If temp >= 0 And temp <= 15 Then shpmain.Width = temp: RaiseEvent Change
    shpmain.Visible = temp > 0
End Property

Public Property Get Power() As Long
    Power = shpmain.Width
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Power = PropBag.ReadProperty("Power", 100)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 150
    UserControl.Width = 300
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Power", Percent, 100
End Sub
