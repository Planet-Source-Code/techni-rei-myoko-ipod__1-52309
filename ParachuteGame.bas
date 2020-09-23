Attribute VB_Name = "ParachuteGame"
Option Explicit
Private Const Pi As Double = 3.14159265358979
Public Type Parachute_Parachuter
    ParaX As Long
    ParaY As Long
    HeliX As Long
    HeliY As Long
    HeliDir As Long
    Released As Boolean
    HeliIsAlive As Boolean
    ParaIsAlive As Boolean
End Type
Public Type Parachute_Bullet
    X As Long
    Y As Long
    Angle As Double
End Type

    Private TurretX As Long, TurretY As Long
    Private TurretLength As Long, TurretAngle As Long
    
    Private Bullets() As Parachute_Bullet
    Private BulletCount As Long, BulletSpeed As Long
    
    Private Parachuters() As Parachute_Parachuter
    Private ParachuteCount As Long, ParachuteSpeed As Long
    
    Private BoardWidth As Long, Lives As Long, Kills As Long
    
Private Sub doDeath()
    If Lives > 0 Then Lives = Lives - 1
End Sub
Public Sub InitParachuteVars()
        TurretAngle = 0
        TurretX = 75
        TurretY = 120
        BoardWidth = TurretX * 2 + 14
        
        ParachuteCount = 4
        ParachuteSpeed = 2
        InitParachuters
        
        BulletSpeed = 5
        Lives = 5
        Kills = 0
End Sub
Public Sub CreateBullet()
        If Lives = 0 Then
            InitParachuteVars
            Exit Sub
        End If
        BulletCount = BulletCount + 1
        ReDim Preserve Bullets(BulletCount)
    With Bullets(BulletCount - 1)
        .Angle = DegreesToRadians(TurretAngle + 180)
        .X = findXY(TurretX + 7, TurretY - 1, 10, .Angle, True)
        .Y = findXY(TurretX + 7, TurretY - 1, 10, .Angle, False)
    End With
End Sub
Public Sub MoveBullets()
    Dim temp As Long, X As Long, Y As Long
        If BulletCount = 0 Then Exit Sub
        For temp = BulletCount - 1 To 0 Step -1
            X = findXY(Bullets(temp).X + 0, Bullets(temp).Y + 0, BulletSpeed + 0, Bullets(temp).Angle + 0, True)
            Y = findXY(Bullets(temp).X + 0, Bullets(temp).Y + 0, BulletSpeed + 0, Bullets(temp).Angle + 0, False)
            If X > 0 And X < BoardWidth And Y > 20 Then
                Bullets(temp).X = X
                Bullets(temp).Y = Y
                CollisionCheck temp
            Else
                DeleteBullet temp
            End If
        Next
End Sub
Private Sub CollisionCheck(index As Long)
    Dim temp As Long
        For temp = 0 To ParachuteCount - 1
            If Parachuters(temp).Released And Parachuters(temp).ParaIsAlive Then
                If IsBulletWithin(index, Parachuters(temp).ParaX, Parachuters(temp).ParaY, 16, 20) Then
                    Parachuters(temp).ParaIsAlive = False
                    Kills = Kills + 1
                End If
            End If
            If Parachuters(temp).HeliIsAlive Then
                If IsBulletWithin(index, Parachuters(temp).HeliX, Parachuters(temp).HeliY, 35, 14) Then
                    Parachuters(temp).HeliIsAlive = False
                    Kills = Kills + 1
                    If Not Parachuters(temp).Released Then
                        Parachuters(temp).ParaIsAlive = False
                        Kills = Kills + 1
                    End If
                End If
            End If
            If Not Parachuters(temp).ParaIsAlive And Not Parachuters(temp).HeliIsAlive Then
                CreateRandomParachuter temp
            End If
        Next
End Sub
Private Function IsWithin(lo As Long, mid As Long, hi As Long) As Boolean
    IsWithin = mid >= lo And mid <= hi
End Function
Private Function IsBulletWithin(index As Long, X As Long, Y As Long, Width As Long, Height As Long) As Boolean
    With Bullets(index)
        IsBulletWithin = IsWithin(X, .X, X + Width - 1) And IsWithin(Y, .Y, Y + Height - 1)
    End With
End Function
Private Sub DeleteBullet(index As Long)
        If BulletCount > 0 Then 'If there will be one or more left afterwards
            If index < BulletCount - 1 Then Bullets(index) = Bullets(BulletCount - 1) 'Switch the one to be deleted with the last one if the one to be deleted isnt the last one
            BulletCount = BulletCount - 1 'then delete the last one
            ReDim Preserve Bullets(BulletCount)
        Else
            BulletCount = 0
            ReDim Bullets(0)
        End If
End Sub
Private Sub DrawBullets(lcdm)
    Dim temp As Long
        If BulletCount = 0 Then Exit Sub
        For temp = 0 To BulletCount
            If Bullets(temp).X > 0 And Bullets(temp).Y > 0 Then lcdm.DrawSquare Bullets(temp).X, Bullets(temp).Y, 2, 2
        Next
End Sub
Private Sub DrawParaTrooper(srcHdc As Long, lcdm, index As Long)
    With Parachuters(index)
        If .HeliIsAlive Then DrawHelicopter srcHdc, lcdm, .HeliX, .HeliY, .HeliDir = -1, .HeliX Mod 10
        If .Released And .ParaIsAlive Then DrawParachuter srcHdc, lcdm, .ParaX, .ParaY
    End With
End Sub
Private Sub InitParachuters()
        ReDim Parachuters(ParachuteCount)
        Dim temp As Long
        For temp = 0 To ParachuteCount - 1
            CreateRandomParachuter temp
        Next
End Sub
Public Sub DrawParachuteScreen(srcHdc As Long, lcdm)
        lcdm.ClearText
        DrawAllTroppers srcHdc, lcdm
        If Lives = 0 Then
            lcdm.PrintText "Game Over", TurretX + 7 - StringWidth("Game Over") / 2, 40
            lcdm.PrintText Kills & " kills", TurretX + 7 - StringWidth(Kills & " kills") / 2, 60
        Else
            TitleBar lcdm, Lives & " lives, & " & Kills & " kills"
            DrawCannon srcHdc, lcdm, TurretX, TurretY, TurretAngle
            DrawBullets lcdm
        End If
End Sub
Private Sub DrawAllTroppers(srcHdc As Long, lcdm)
    Dim temp As Long
        For temp = 0 To ParachuteCount - 1
            DrawParaTrooper srcHdc, lcdm, temp
        Next
End Sub
Public Sub MoveTurret(direction As Long)
        TurretAngle = TurretAngle + direction
        If TurretAngle > 90 And TurretAngle < 180 Then TurretAngle = 90
        If TurretAngle < 0 Then TurretAngle = 360 + TurretAngle
        If TurretAngle < 270 And TurretAngle > 180 Then TurretAngle = 270
        If TurretAngle > 359 Then TurretAngle = TurretAngle - 360
End Sub
Public Sub MoveParachuters()
    Dim temp As Long
        For temp = 0 To ParachuteCount
            Parachuters(temp).HeliX = Parachuters(temp).HeliX + ParachuteSpeed * Parachuters(temp).HeliDir
            If Parachuters(temp).Released Then
                Parachuters(temp).ParaY = Parachuters(temp).ParaY + ParachuteSpeed
            Else
                If Parachuters(temp).HeliIsAlive Then
                    If Parachuters(temp).HeliDir > 0 Then
                        If Parachuters(temp).ParaX < Parachuters(temp).HeliX Then Parachuters(temp).Released = True
                    Else
                        If Parachuters(temp).ParaX > Parachuters(temp).HeliX Then Parachuters(temp).Released = True
                    End If
                End If
            End If
        Next
        For temp = ParachuteCount - 1 To 0 Step -1
            If Parachuters(temp).ParaY >= TurretY + 15 Then 'If parachuter is off screen
                If Parachuters(temp).HeliX <= -35 Or Parachuters(temp).HeliX > BoardWidth Then 'if helicopter is off screen
                    If Parachuters(temp).ParaIsAlive Then
                        doDeath
                    End If
                    CreateRandomParachuter temp
                End If
            End If
        Next
End Sub
Private Sub CreateRandomParachuter(index As Long)
    Dim X As Long, Y As Long, direction As Long
        Randomize Timer
        X = Rnd * (BoardWidth - 16)
        Randomize Timer
        Y = 20 + Rnd * 25
        direction = 1
        Randomize Timer
        If Rnd < 0.5 Then direction = -1
        NewParachuter index, X, Y, direction
End Sub
Private Sub NewParachuter(index As Long, X As Long, Y As Long, direction As Long)
    With Parachuters(index)
        .HeliDir = direction
        If direction = -1 Then .HeliX = BoardWidth Else .HeliX = -35
        .ParaX = X
        .ParaY = Y
        .HeliY = Y
        .HeliIsAlive = True
        .ParaIsAlive = True
        .Released = False
    End With
End Sub
Private Function findXY(X As Single, Y As Single, Distance As Single, Angle As Double, Optional isx As Boolean = True) As Single
    If isx = True Then findXY = X + Sin(Angle) * Distance Else findXY = Y + Cos(Angle) * Distance
End Function
Private Function rad2deg(Radians As Double) As Double
    rad2deg = Radians * 180
End Function
Private Function DegreesToRadians(ByVal Degrees As Double) As Double 'Converts Degrees to Radians.
    DegreesToRadians = Degrees * (Pi / 180)
End Function
Private Function DrawHelicopter(srcHdc As Long, lcdm, X As Long, Y As Long, direction As Boolean, Bladewidth As Long)
    Dim top As Long, leftblade As Long, rightblade As Long, otherblade As Long
    otherblade = 32 'top=6
    leftblade = 9
    If Not direction Then
        otherblade = 2
        top = 14
        leftblade = 21
    End If
    rightblade = leftblade + 4
    TransBLT srcHdc, 0, top, srcHdc, 35, top, 35, 14, lcdm.hdc, X, Y
    lcdm.DrawLine X + leftblade, Y, -Bladewidth, 1
    lcdm.DrawLine X + rightblade, Y, Bladewidth, 1
End Function
Private Sub DrawParachuter(srcHdc As Long, lcdm, X As Long, Y As Long)
    TransBLT srcHdc, 0, 28, srcHdc, 35, 28, 16, 20, lcdm.hdc, X, Y
End Sub
Private Sub DrawCannon(srcHdc As Long, lcdm, ByVal X As Long, ByVal Y As Long, Angle As Long)
    Dim temp As Double, lefts As Long, top As Long
    TransBLT srcHdc, 17, 33, srcHdc, 52, 33, 14, 15, lcdm.hdc, X, Y
    X = X + 7
    temp = DegreesToRadians(Angle)
    lefts = findXY(X + 0, Y + 0, 10, temp)
    top = findXY(X + 0, Y + 0, 10, temp, False)
    lcdm.DrawLine X, Y, X - lefts + 1, Y - top + 1
    lcdm.DrawLine X - 1, Y, X - lefts + 1, Y - top + 1
End Sub
