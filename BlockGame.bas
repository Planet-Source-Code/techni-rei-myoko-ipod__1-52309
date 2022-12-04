Attribute VB_Name = "BlockGame"
Option Explicit
Public Const AngleFactor As Long = 1.15
    'Removed the block type
    'Grid Details
    Private TileWidth As Long, TileHeight As Long
    Private GridWidth As Long, GridHeight As Long
    Private GridTop As Long, GridLeft As Long
    Private GridWhiteSpace As Long
    Private Grid() As Boolean
    
    'Variables
    Private Total As Long, Lives As Long, InGame As Boolean
    
    'Paddle Details
    Private PaddleTop As Long, PaddleLeft As Long
    Private PaddleHeight As Long, PaddleWidth As Long
    
    'Puck Details
    Private PuckTop As Long, PuckLeft As Long
    Private PuckDirX As Long, PuckDirY As Long
    Private PuckSize As Long

Private lcdm
Private Sub DrawTitleBar()
    lcdm.DrawSquare 0, 0, lcdm.Width, 20, lcdm.BackColor, True
    TitleBar lcdm, Lives & " lives, & " & Total & " blocks", PlayPause
End Sub
Private Sub BlockDeath()
    Lives = Lives - 1
    DrawTitleBar
    InGame = False
    
    If Lives = 0 Then
        lcdm.ClearText
        lcdm.PrintText "Game Over", (lcdm.Width - StringWidth("Game Over")) / 2, (lcdm.Height - StringHeight("Game Over")) / 2, False
    Else
        InitializePuck 1, 68, 1, 1, 3
        DrawScreen
    End If
End Sub
Public Sub StartGame()
    If Lives = 0 Then
        InitializeDefaults lcdm
        DrawPaddle
    Else
        InGame = True
        DrawPaddle
    End If
End Sub
Public Sub InitializeDefaults(Screen)
    Set lcdm = Screen
    lcdm.ClearText
    InitializeGrid 10, 4, 13, 4, 24, 5, 2
    InitializePaddle 20, 2, 125, 0
    InitializePuck 1, 68, 1, 1, 3
    Lives = 3
    ResetBlockGame 68
    DrawTitleBar
    DrawScreen
End Sub
Private Sub InitializeGrid(TWidth As Long, THeight As Long, GWidth As Long, GHeight As Long, GTop As Long, GLeft As Long, GWhiteSpace As Long)
        TileWidth = TWidth
        TileHeight = THeight
        GridWidth = GWidth
        GridHeight = GHeight
        GridTop = GTop
        GridLeft = GLeft
        GridWhiteSpace = GWhiteSpace
        ReDim Grid(1 To GridHeight, 1 To GridWidth)
        Total = GridWidth * GridHeight
End Sub
Private Sub InitializePaddle(Width As Long, Height As Long, top As Long, Optional Left As Long)
        PaddleTop = top
        PaddleHeight = Height
        PaddleLeft = Left
        PaddleWidth = Width
End Sub
Private Sub InitializePuck(Left As Long, top As Long, dirX As Long, dirY As Long, Size As Long)
        PuckLeft = Left
        PuckTop = top
        PuckDirX = dirX
        PuckDirY = dirY
        PuckSize = Size
End Sub

Public Sub ResetBlockGame(Optional PuckTop As Long = 68, Optional GridValue As Boolean = True)
    Dim temp As Long, temp2 As Long
    For temp = 1 To GridHeight
        For temp2 = 1 To GridWidth
            Grid(temp, temp2) = GridValue
        Next
    Next
    
    PaddleLeft = 0
    InitializePuck 1, PuckTop, 1, 1, PuckSize
End Sub
Private Sub DrawTile(Col As Long, Row As Long, Optional color As Long = vbBlack)
    Dim Left As Long, top As Long
        If Grid(Row, Col) Then
            Left = getLeft(Col)
            top = GetTop(Row)
            lcdm.DrawSquare Left, top, TileWidth, TileHeight, color, True
        End If
End Sub
Private Sub DrawPaddle(Optional color As Long = vbBlack)
        lcdm.DrawSquare PaddleLeft, PaddleTop, PaddleWidth, PaddleHeight, color, True
End Sub
Private Sub DrawPuck(Optional color As Long = vbBlack)
        lcdm.DrawSquare PuckLeft, PuckTop, PuckSize, PuckSize, color, True
End Sub
Private Sub DrawGrid()
    Dim temp As Long, temp2 As Long
    For temp = 1 To GridHeight
        For temp2 = 1 To GridWidth
            DrawTile temp2, temp
        Next
    Next
End Sub
Public Sub DrawScreen()
    DrawPuck
    DrawPaddle
    DrawGrid
End Sub
Public Sub MovePaddle(value As Long, Optional dir As Long = 1)
    ClearPaddle
        Dim wid As Long
        wid = lcdm.Width
        PaddleLeft = PaddleLeft + value * dir
        If PaddleLeft < 0 Then PaddleLeft = 0
        If PaddleLeft + PaddleWidth > wid Then PaddleLeft = wid - PaddleWidth
    DrawPaddle
End Sub
Private Sub ClearPuck()
    DrawPuck lcdm.BackColor
End Sub
Private Sub ClearPaddle()
    DrawPaddle lcdm.BackColor
End Sub
Public Sub MovePuck()
    If Not InGame Then Exit Sub
    ClearPuck
        If Not InGame Then Exit Sub
        PuckLeft = PuckLeft + PuckDirX
        PuckTop = PuckTop + PuckDirY
        
        If PuckTop = 20 Then PuckDirY = 1
        If PuckTop + PuckSize = lcdm.Height - 1 Then BlockDeath 'lose a point
        If PuckLeft = 0 Then PuckDirX = 1
        If PuckLeft + PuckSize = lcdm.Width - 1 Then PuckDirX = -1
        
        If PuckTop + PuckSize >= PaddleTop And PuckTop <= PaddleTop + PaddleHeight - 1 Then
            If PuckLeft + PuckSize >= PaddleLeft And PuckLeft <= PaddleLeft + PaddleWidth Then PuckDirY = -1
        End If
        
        CollisionDetect PuckLeft, PuckTop
        CollisionDetect PuckLeft + PuckSize - 1, PuckTop
        CollisionDetect PuckLeft, PuckTop + PuckSize - 1
        CollisionDetect PuckLeft + PuckSize - 1, PuckTop + PuckSize - 1
    DrawPuck
End Sub
Private Sub CollisionDetect(X As Long, Y As Long)
        Dim trow As Long, tcol As Long, tX As Long, tY As Long
        trow = Row(Y)
        tcol = Col(X)
            If trow >= 1 And trow <= GridHeight Then
                If tcol >= 1 And tcol <= GridWidth Then
                    If Grid(trow, tcol) = True Then
                        tX = getLeft(tcol)
                        tY = GetTop(trow)
                        If PuckTop + PuckSize >= tY And PuckTop <= tY + TileHeight - 1 Then
                            If PuckLeft + PuckSize >= tX And PuckLeft <= tX + TileWidth Then
                                DrawTile tcol, trow, &HC8DDC1
                                Grid(trow, tcol) = False
                                Total = Total - 1
                                DrawTitleBar
                                SwitchX
                                SwitchY
                                If Total = 0 Then BlockWin
                            End If
                        End If
                    End If
                End If
            End If
End Sub
Private Sub BlockWin()
    Const youwin = "You Win"
    InGame = False
    lcdm.ClearText
    lcdm.PrintText youwin, (lcdm.Width - StringWidth(youwin)) / 2, (lcdm.Height - StringHeight(youwin)) / 2, False
End Sub
Private Sub SwitchX()
    PuckDirX = SwithDirection(PuckDirX)
End Sub
Private Sub SwitchY()
    PuckDirY = SwithDirection(PuckDirY)
End Sub

Private Function SwithDirection(orig As Long) As Long
    SwithDirection = IIf(orig = -1, 1, -1)
End Function

Private Function Col(ByVal X As Long) As Long
        X = (X - GridLeft) \ (TileWidth + GridWhiteSpace)
        Col = X + 1
End Function
Private Function Row(ByVal Y As Long) As Long
        Y = (Y - GridTop) \ (TileHeight + GridWhiteSpace)
        Row = Y + 1
End Function
Private Function getLeft(Col As Long) As Long
        getLeft = (Col - 1) * (TileWidth + GridWhiteSpace) + GridLeft
End Function
Private Function GetTop(Row As Long) As Long
        GetTop = (Row - 1) * (TileHeight + GridWhiteSpace) + GridTop
End Function
