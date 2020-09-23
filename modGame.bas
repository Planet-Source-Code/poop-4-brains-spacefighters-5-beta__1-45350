Attribute VB_Name = "modGame"
Option Explicit

Public Const PI = 3.1459

Type Controls
    Forward As Integer
    Left As Integer
    Right As Integer
    Shoot As Integer
End Type

Type Player
    x As Long
    y As Long
    Xs As Long
    Ys As Long
    Rot As Long
    Act As Long

    Score As Long
    Name As Long
    Ammo As Long
    HP As Long
    MHP As Long
    Ty As Long
    Reload As Long
    Tag As Long
End Type

Type Asteroid
    x As Long
    y As Long
    Xs As Long
    Ys As Long
    Rot As Long
    Act As Long
    Ty As Long
End Type

Type Shot
    x As Long
    y As Long
    Xs As Long
    Ys As Long
    Act As Long
    Tag As Long
End Type

Public Speed As Long
Public Running As Boolean
Public Paused As Boolean
Public P(1 To 4) As Player
Public A(1 To 10) As Asteroid
Public S(1 To 15) As Shot
Public Pc(1 To 4) As Controls

Function RotShip(Dir As Integer, Pl As Player)
Select Case Dir
Case 0 ' Left
    Pl.Rot = Pl.Rot - 1: If Pl.Rot < 0 Then Pl.Rot = 7
Case 1 ' Right
    Pl.Rot = Pl.Rot + 1: If Pl.Rot > 7 Then Pl.Rot = 0
End Select
End Function

Function MoveObjects()
MoveShips
MoveShots
MoveAst
End Function

Function MoveAst()
Dim i As Long

For i = 1 To 10
    If A(i).Act = True Then
        A(i).x = A(i).x + A(i).Xs
        A(i).y = A(i).y + A(i).Ys
        A(i).Rot = A(i).Rot + 1: If A(i).Rot > 9 Then A(i).Rot = 0

        If A(i).x + 50 < 0 Or A(i).y + 45 < 0 Or A(i).y > frmGame.Board.ScaleHeight Or A(i).x > frmGame.Board.ScaleWidth Then A(i).Act = False
    End If
Next i
End Function

Function MoveShips()
Dim i As Long

For i = 1 To 4
    If P(i).Act = True Then
        P(i).Reload = P(i).Reload - 1: If P(i).Reload < 0 Then P(i).Reload = 0
        P(i).x = P(i).x + P(i).Xs
        P(i).y = P(i).y + P(i).Ys

        P(i).Xs = P(i).Xs * 0.8
        P(i).Ys = P(i).Ys * 0.8

        If P(i).x < -50 Then P(i).x = frmGame.Board.ScaleWidth
        If P(i).y < -45 Then P(i).y = frmGame.Board.ScaleHeight
        If P(i).y > frmGame.Board.ScaleHeight Then P(i).y = -40
        If P(i).x > frmGame.Board.ScaleWidth Then P(i).x = -45
    End If
Next i
End Function

Function DoAllKeys()
Dim O As Integer

For O = 1 To 4
    If P(O).Act = True Then
        DoKeys P(O), Pc(O).Forward, Pc(O).Left, Pc(O).Right, Pc(O).Shoot
    End If
Next O
End Function

Function MoveShots()
Dim i As Long, r As Long

For i = 1 To 15
    If S(i).Act = True Then
        S(i).x = S(i).x + S(i).Xs
        S(i).y = S(i).y + S(i).Ys

        If S(i).x + 15 < 0 Or S(i).y + 15 < 0 Or S(i).y > frmGame.Board.ScaleHeight Or S(i).x > frmGame.Board.ScaleWidth Then S(i).Act = False
        
        For r = 1 To 4
            If r <> S(i).Tag Then
                If CollisionDetect(S(i).x, S(i).y, 15, 15, 0, 0, frmGfx.ShM.hdc, P(r).x, P(r).y, 50, 45, P(r).Rot * 50, 0, frmGfx.ShipM(P(r).Ty).hdc) = True Then
                    S(i).Act = False
                    Hurt P(r), 5
                    Exit For
                End If
            End If
        Next r

    End If
Next i
End Function

Function Hurt(Pl As Player, Dmg As Long)
Pl.HP = Pl.HP - Dmg

If Pl.HP <= 0 Then Pl.Act = False
End Function

Function ShootBullet(Pl As Player)
Dim w As Long, Ang As Long

If Pl.Reload > 0 Then Exit Function

For w = 1 To 15
    If S(w).Act = False Then
        S(w).Act = True

        Ang = Pl.Rot * 45 - 90
        S(w).x = Pl.x + 25 + Cos(Ang * PI / 180) * 10
        S(w).y = Pl.y + (45 / 2) + Sin(Ang * PI / 180) * 10

        S(w).Xs = Cos(Ang * PI / 180) * 15
        S(w).Ys = Sin(Ang * PI / 180) * 15
        S(w).Tag = Pl.Tag
        Pl.Reload = 5
        Exit For
    End If
Next w
End Function
Function DoKeys(Pl As Player, Forward, Left, Right, Shoot)
If GetAsyncKeyState(Forward) Then
    Dim Ang
    Ang = Pl.Rot * 45 - 90
    Pl.Xs = Cos(Ang * PI / 180) * 10
    Pl.Ys = Sin(Ang * PI / 180) * 10
End If

If GetAsyncKeyState(Left) Then
    RotShip 0, Pl
End If

If GetAsyncKeyState(Right) Then
    RotShip 1, Pl
End If

If GetAsyncKeyState(Shoot) Then
    ShootBullet Pl
End If
End Function


