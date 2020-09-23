Attribute VB_Name = "ModEngine"
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Public Const sRed = 0
Public Const sBlue = 1
Public Const sGreen = 2
Public Const sMine = 3

Public Const pShield = 0

Type Player
'movement and placement info
x As Long
y As Long
XS As Long
YS As Long
Boost As Long
Rotate As Long

'identifying properties
Score As Long
Band As Long
Nick As String
Health As Long
MAXHP As Long
id As Long
Type As Long
Act As Long
Reload As Long
Armor As Long
ReSmoke As Long

'powerups
Shield As Long
DeathBall As Long
Mines As Long
Superman As Long

'AI Info
Target As Long 'which player to target
AI As Boolean 'determines if its a comp player
BoostLeft As Long
End Type

Type Shot
'movement and placement info
x As Long
y As Long
XS As Long
YS As Long
Rotate As Long

'identifying info
Act As Boolean
id As Long
Type As Long
Damage As Long 'how much damage the bullet does
Count As Long 'how long it lasts
End Type

Type PowerUp
'placement and animation info
x As Long
y As Long

'identifying info
Count As Long 'how long it has left (powerups dont stay forever ya know)
Act As Long
Type As Long 'type of powerup
End Type

Type Explosion
'movement and placement info
x As Long
y As Long
XS As Long
YS As Long

'anim
Frame As Long 'current animation frame
Frames As Long 'how many frames are in animation

'identifying info
Act As Long
Type As Long
End Type

Type text
Major As String 'title
Minor As String 'subtitle
End Type

Type GameInfo ' not used
VictoryType As Long 'what kind of victory

ReqScore As Long 'the required score of a player to win (if any)
Time As Long 'time left in the game (if a timed game)
End Type

Type Asteroid
x As Long
y As Long
Rotate As Long
XS As Long
YS As Long
Act As Long
End Type

Public P(1 To 4) As Player
Public S() As Shot
Public PUP(20) As PowerUp
Public Explo(1 To 30) As Explosion
Public A(10) As Asteroid
Public Message As text
Public Running As Boolean
Public Bands As Boolean
Public Speed As Long
'masking hdcs for collison detection
Public PMask(5) As Long, SMask(10) As Long, PUPMask(10) As Long, AMask As Long
Public Playing As Boolean
Public TeamName(1 To 2) As String
Public i As Long

Function MakePowerup(T As Long)
Dim i As Long
For i = 0 To 20
If PUP(i).Act = False Then
PUP(i).Type = T
PUP(i).x = Int(Rnd * frmGame.board.ScaleWidth)
PUP(i).y = Int(Rnd * frmGame.board.ScaleHeight)
PUP(i).Count = 120
PUP(i).Act = True
Debug.Print "Powerup MadE"
Exit For
End If
Next i
End Function

Function MakePlayer(x, y, Name, Band, PType) As Boolean
Dim c As Long 'counter variable

For c = 1 To 4
If P(c).Act = False Then
'set the coors
P(c).x = x
P(c).y = y

'reset the props
P(c).XS = 0
P(c).YS = 0
P(c).Rotate = 0

'set other props
P(c).id = c
P(c).Type = PType
P(c).Nick = Name

'activate player
P(c).Act = True

P(c).Health = 50
Select Case PType
Case sRed
P(c).Armor = 2
Case sBlue
P(c).Armor = 0
Case sGreen
P(c).Armor = 4
End Select

P(c).Shield = 30
Exit For
End If
Next c
End Function

Function TurnLeft(P As Player)
P.Rotate = P.Rotate - 1: If P.Rotate < 0 Then P.Rotate = 7
End Function

Function TurnRight(P As Player)
P.Rotate = P.Rotate + 1: If P.Rotate > 7 Then P.Rotate = 0
End Function

Function MoveForward(Player As Player)
Dim Acc As Long
Player.Boost = 1
Select Case Player.Type
Case sBlue
Acc = 9
Case sRed
Acc = 7
Case sGreen
Acc = 5
End Select
If Player.Rotate = 0 Then: SetSpeed 0, -Acc, Player
If Player.Rotate = 1 Then: SetSpeed Acc, -Acc, Player
If Player.Rotate = 2 Then: SetSpeed Acc, 0, Player
If Player.Rotate = 3 Then: SetSpeed Acc, Acc, Player
If Player.Rotate = 4 Then: SetSpeed 0, Acc, Player
If Player.Rotate = 5 Then: SetSpeed -Acc, Acc, Player
If Player.Rotate = 6 Then: SetSpeed -Acc, 0, Player
If Player.Rotate = 7 Then: SetSpeed -Acc, -Acc, Player

Select Case Player.Type
Case sBlue
If Player.XS < -20 Then Player.XS = -20
If Player.XS > 20 Then Player.XS = 20
If Player.YS < -20 Then Player.YS = -20
If Player.YS > 20 Then Player.YS = 20
Case sRed
If Player.XS < -15 Then Player.XS = -15
If Player.XS > 15 Then Player.XS = 15
If Player.YS < -15 Then Player.YS = -15
If Player.YS > 15 Then Player.YS = 15
Case sGreen
If Player.XS < -10 Then Player.XS = -10
If Player.XS > 10 Then Player.XS = 10
If Player.YS < -10 Then Player.YS = -10
If Player.YS > 10 Then Player.YS = 10
End Select
End Function

Function Fire(Player As Player, T As Long)
If Player.Reload > 0 Then Exit Function
Dim c As Long
For c = 0 To UBound(S())
If S(c).Act = False Then
S(c).x = Player.x + 37 - 15
S(c).y = Player.y + 37 - 15
S(c).Count = 20
S(c).id = Player.id
Player.Reload = 3
If Player.Rotate = 0 Then: S(c).XS = 0
If Player.Rotate = 1 Then: S(c).XS = 14
If Player.Rotate = 2 Then: S(c).XS = 14
If Player.Rotate = 3 Then: S(c).XS = 14
If Player.Rotate = 4 Then: S(c).XS = 0
If Player.Rotate = 5 Then: S(c).XS = -14
If Player.Rotate = 6 Then: S(c).XS = -14
If Player.Rotate = 7 Then: S(c).XS = -14
If Player.Rotate = 0 Then: S(c).YS = -14
If Player.Rotate = 1 Then: S(c).YS = -14
If Player.Rotate = 2 Then: S(c).YS = 0
If Player.Rotate = 3 Then: S(c).YS = 14
If Player.Rotate = 4 Then: S(c).YS = 14
If Player.Rotate = 5 Then: S(c).YS = 14
If Player.Rotate = 6 Then: S(c).YS = 0
If Player.Rotate = 7 Then: S(c).YS = -14

S(c).Type = T
S(c).Act = True
Select Case T
Case sBlue
S(c).Damage = 1
S(c).Rotate = 0
Case sRed
S(c).Damage = 2
S(c).Rotate = Player.Rotate
Case sGreen
S(c).Damage = 3
S(c).Rotate = 0
Case sMine
S(c).Damage = 10
S(c).Rotate = 0
S(c).XS = 0
S(c).YS = 0
S(c).Count = 500
Player.Mines = Player.Mines - 1
End Select

If Player.Superman > 0 Then
S(c).Damage = S(c).Damage * 1.5
S(c).Count = S(c).Count + 20
S(c).Type = S(c).Type + 4
End If
Exit For
End If
Next c
End Function

Function SetSpeed(XSpeed, YSpeed, Player As Player)
Player.XS = Player.XS + XSpeed
Player.YS = Player.YS + YSpeed
End Function

Function MovePlayer(Player As Player)
On Error Resume Next
Dim c As Long, Friction As Long, SmokeThing, i, AV

For c = 0 To 20
If PUP(c).Act = True Then
If CollisionDetect(Player.x, Player.y, 75, 75, Player.Rotate * 75, 0, PMask(Player.Type), PUP(c).x, PUP(c).y, 30, 30, 0, 0, PUPMask(PUP(c).Type)) = True Then
PUP(c).Act = False
Select Case PUP(c).Type
Case 0
Player.Shield = Player.Shield + 90
Case 1
Player.Mines = Player.Mines + 5
Case 2
Do
Player.x = Int(Rnd * frmGame.board.ScaleWidth) - 37
Player.y = Int(Rnd * frmGame.board.ScaleHeight) - 37
For i = 1 To 4
If P(i).id <> Player.id And CollisionDetect(Player.x, Player.y, 75, 75, Player.Rotate * 75, 0, PMask(Player.Type), P(i).x, P(i).y, 75, 75, P(i).Rotate * 75, 0, PMask(P(i).Type)) = False Then Exit Do
Next i
DoEvents
Loop
Case 3
Player.Superman = Player.Superman + 150
Player.Armor = Player.Armor + 2
End Select
End If
End If
Next c

If Player.Health < (Player.MAXHP * 0.7) Then
SmokeThing = Player.Health \ 2 + (Player.Health \ 4)
Player.ReSmoke = Player.ReSmoke + 1
If Player.ReSmoke >= SmokeThing Then
DoExplo Player.x + 22, Player.y + 22, IIf(Int(Rnd * 2) = 0, Int(Rnd * 5) + 1, -(Int(Rnd * 5) + 1)), IIf(Int(Rnd * 2) = 0, Int(Rnd * 5) + 1, -(Int(Rnd * 5) + 1)), 5, 4
Player.ReSmoke = 0
End If
End If

For c = 0 To UBound(S())
If S(c).Act = True Then
If CollisionDetect(Player.x, Player.y, 75, 75, Player.Rotate * 75, 0, PMask(Player.Type), S(c).x, S(c).y, 30, 30, 0, 0, SMask(S(c).Type)) = True And S(c).id <> Player.id Then
If Player.Shield <= 0 Then Player.Health = Player.Health - S(c).Damage \ (Player.Armor + 1)

S(c).Act = False:
Select Case S(c).Type
Case sRed
DoExplo S(c).x, S(c).y, 0, 0, 0, 5
Case sBlue
DoExplo S(c).x, S(c).y, 0, 0, 2, 3
Case sGreen
DoExplo S(c).x, S(c).y, 0, 0, 3, 4
Case sMine
DoExplo S(c).x, S(c).y, 0, 0, 7, 10
Case 4
DoExplo S(c).x, S(c).y, 0, 0, 0, 5
Case 5
DoExplo S(c).x, S(c).y, 0, 0, 2, 3
Case 6
DoExplo S(c).x, S(c).y, 0, 0, 3, 4
End Select

Player.Armor = Player.Armor - 1: If Player.Armor < 0 Then Player.Armor = 0
If Bands = True And Player.Band <> GetIDBand(S(c).id) Then Player.Target = S(c).id
If Player.Health <= 0 Then
Player.Act = False

Select Case Player.Type
Case sGreen
DoExplo Player.x + 20, Player.y + 20, 0, 0, 6, 6
Case Else
DoExplo Player.x + 20, Player.y + 20, 0, 0, 1, 6
End Select
End If
End If
End If
Next c

Select Case Player.Type
Case 0
Friction = 2
Case 1
Friction = 3
Case 2
Friction = 1
End Select

Select Case Player.XS
Case Is < 0
Player.XS = Player.XS + Friction
Case Is > 0
Player.XS = Player.XS - Friction
End Select
Select Case Player.YS
Case Is < 0
Player.YS = Player.YS + Friction
Case Is > 0
Player.YS = Player.YS - Friction
End Select

Player.x = Player.x + Player.XS
For i = 1 To 4
If P(i).id <> Player.id And CollisionDetect(Player.x, Player.y, 75, 75, Player.Rotate * 75, 0, PMask(Player.Type), P(i).x, P(i).y, 75, 75, P(i).Rotate * 75, 0, PMask(P(i).Type)) = True Then
Player.x = Player.x - Player.XS
P(i).x = P(i).x - P(i).XS
Player.XS = -(Player.XS)
P(i).XS = -(P(i).XS)
End If
Next i

Player.y = Player.y + Player.YS
For i = 1 To 4
If P(i).Act = True And P(i).id <> Player.id And CollisionDetect(Player.x, Player.y, 75, 75, Player.Rotate * 75, 0, PMask(Player.Type), P(i).x, P(i).y, 75, 75, P(i).Rotate * 75, 0, PMask(P(i).Type)) = True Then
Player.y = Player.y - Player.YS
Player.YS = -(Player.YS)
P(i).y = P(i).y - P(i).YS
P(i).YS = -(P(i).YS)
End If
Next i

With frmGame
If Player.x < -37 Then Player.x = .board.ScaleWidth + 37
If Player.y < -37 Then Player.y = .board.ScaleHeight + 37
If Player.y > .board.ScaleHeight + 37 Then Player.y = -37
If Player.x > .board.ScaleWidth + 37 Then Player.x = -37
End With

For c = 0 To 10
If CollisionDetect(A(c).x, A(c).y, 30, 30, A(c).Rotate * 45, 0, AMask, Player.x, Player.y, 75, 75, Player.Rotate * 75, 0, PMask(Player.Type)) = True And A(c).Act = True Then

If Player.Shield <= 0 Then
Player.Health = Player.Health - 2

If Player.Health < 0 Then
Player.Act = False
Select Case Player.Type
Case sGreen
DoExplo Player.x + 20, Player.y + 20, 0, 0, 6, 6
Case Else
DoExplo Player.x + 20, Player.y + 20, 0, 0, 1, 6
End Select
End If
End If
A(c).Act = False

DoExplo A(c).x + 5, A(c).y + 5, IIf(Int(Rnd * 2) = 0, Int(Rnd * 5) + 1, -(Int(Rnd * 5) + 1)), IIf(Int(Rnd * 2) = 0, Int(Rnd * 5) + 1, -(Int(Rnd * 5) + 1)), 5, 4
DoExplo A(c).x + 5, A(c).y + 5, IIf(Int(Rnd * 2) = 0, Int(Rnd * 5) + 1, -(Int(Rnd * 5) + 1)), IIf(Int(Rnd * 2) = 0, Int(Rnd * 5) + 1, -(Int(Rnd * 5) + 1)), 5, 4
End If
Next c
Player.Reload = Player.Reload - 1: If Player.Reload < 0 Then Player.Reload = 0
Player.Boost = Player.Boost - 1: If Player.Boost < 0 Then Player.Boost = 0
Player.Shield = Player.Shield - 1: If Player.Shield < 0 Then Player.Shield = 0
Player.Superman = Player.Superman - 1: If Player.Superman < 0 Then Player.Superman = 0
End Function

Function MoveShots()
Dim c As Long

For c = 0 To 20
If S(c).Act = True Then
If S(c).Count <= 0 Then S(c).Act = False
S(c).Count = S(c).Count - 1
S(c).x = S(c).x + S(c).XS
S(c).y = S(c).y + S(c).YS

With frmGame
If S(c).x < -30 Then S(c).x = .board.ScaleWidth + 30
If S(c).y < -30 Then S(c).y = .board.ScaleHeight + 30
If S(c).y > .board.ScaleHeight + 30 Then S(c).y = -30
If S(c).x > .board.ScaleWidth + 30 Then S(c).x = -30
End With
End If
Next c

For c = 1 To 30
If Explo(c).Act = True Then
Explo(c).x = Explo(c).x + Explo(c).XS
Explo(c).y = Explo(c).y + Explo(c).YS
Explo(c).Frame = Explo(c).Frame + 1: If Explo(c).Frame > Explo(c).Frames Then Explo(c).Act = False
End If
Next c


'move asteroids
Dim C3 As Long
For c = 0 To 10
If A(c).Act = True Then
A(c).x = A(c).x + A(c).XS
A(c).y = A(c).y + A(c).YS
A(c).Rotate = A(c).Rotate + 1: If A(c).Rotate > 35 Then A(c).Rotate = 0

With frmGame
If A(c).x < -30 Then A(c).x = .board.ScaleWidth + 30
If A(c).x > .board.ScaleWidth + 5 Then A(c).x = -30
If A(c).y < -30 Then A(c).y = .board.ScaleHeight + 30
If A(c).y > .board.ScaleHeight + 30 Then A(c).y = -30
End With

For C3 = 0 To UBound(S())
If CollisionDetect(A(c).x, A(c).y, 30, 30, A(c).Rotate * 45, 0, AMask, S(C3).x, S(C3).y, 30, 30, 0, 0, SMask(S(C3).Type)) = True Then
S(C3).Act = False
Select Case S(c).Type
Case sRed
DoExplo S(c).x, S(c).y, 0, 0, 0, 5
Case sBlue
DoExplo S(c).x, S(c).y, 0, 0, 2, 3
Case sGreen
DoExplo S(c).x, S(c).y, 0, 0, 3, 4
End Select

A(c).Act = False

DoExplo A(c).x + 5, A(c).y + 5, IIf(Int(Rnd * 2) = 0, Int(Rnd * 5) + 1, -(Int(Rnd * 5) + 1)), IIf(Int(Rnd * 2) = 0, Int(Rnd * 5) + 1, -(Int(Rnd * 5) + 1)), 5, 4
DoExplo A(c).x + 5, A(c).y + 5, IIf(Int(Rnd * 2) = 0, Int(Rnd * 5) + 1, -(Int(Rnd * 5) + 1)), IIf(Int(Rnd * 2) = 0, Int(Rnd * 5) + 1, -(Int(Rnd * 5) + 1)), 5, 4
End If
Next C3
End If
Next c

For i = 0 To 20
If PUP(i).Act = True Then
PUP(i).Count = PUP(i).Count - 1: If PUP(i).Count < 0 Then PUP(i).Act = False
End If
Next i
End Function

Function DoKeys(Player As Player, L, R, U, Shoot, Special)
If GetAsyncKeyState(L) Then TurnLeft Player
If GetAsyncKeyState(R) Then TurnRight Player
If GetAsyncKeyState(U) Then MoveForward Player
If GetAsyncKeyState(Shoot) Then Fire Player, Player.Type
If GetAsyncKeyState(Special) Then
If Player.Mines > 0 Then
Fire Player, 3
End If
End If
End Function

Function DoExplo(x, y, XSp, YSp, eType, Frames)
Dim c As Long
For c = 1 To 30
If Explo(c).Act = False Then
Explo(c).Frame = 0
Explo(c).Act = True
Explo(c).x = x
Explo(c).y = y
Explo(c).XS = XSp
Explo(c).YS = YSp
Explo(c).Type = eType
Explo(c).Frames = Frames
Exit For
End If
Next c
End Function

Function DoAI(Player As Player)
Dim Try As Long
If Player.Target = 0 Then
Try = 0
Do
Player.Target = Int(Rnd * 4) + 1
If Player.Target <> Player.id And P(Player.Target).Act = True And Player.Band <> P(Player.Target).Band Then
Exit Do
Else
Try = Try + 1: If Try > 8 Then Exit Do
DoEvents
End If
Loop
End If

If Int(Rnd * 10) = Int(Rnd * 9) Then
For i = 0 To UBound(S())
If S(i).Act = True Then
If S(i).x + 15 > Player.x - 75 And S(i).x + 15 < Player.x + (75 * 2) And S(i).y + 15 > Player.y - 75 And S(i).y + 15 < Player.y + (75 * 2) Then
Player.Rotate = FindAng(Player.x + 37, Player.y + 37, S(i).x + 15, S(i).y + 15) - 2
If Player.Rotate < 0 Then Player.Rotate = 7
If Player.Rotate > 7 Then Player.Rotate = 0
MoveForward Player
End If
End If
Next i
End If

If Int(Rnd * 5) = Int(Rnd * 6) Then 'seek and destroy
Player.Rotate = FindAng(Player.x + 37, Player.y + 37, P(Player.Target).x + 37, P(Player.Target).y + 37)
MoveForward Player
Fire Player, Player.Type
End If

If Player.Mines > 0 And Int(Rnd * 5) = Int(Rnd * 4) Then Fire Player, 3
If Player.BoostLeft > 0 Then
MoveForward Player
Player.BoostLeft = Player.BoostLeft - 1
End If
End Function

Function FindAng(SrcX, SrcY, DstX, DstY) As Long
Dim XS, YS, Range
Range = 40
Select Case DstX
Case Is < SrcX - Range
XS = -1
Case Is > SrcX + Range
XS = 1
Case Is < SrcX + Range And DstX > SrcX - Range
XS = 0
End Select

Select Case DstY
Case Is < SrcY - Range
YS = -1
Case Is > SrcY + Range
YS = 1
Case Is < SrcY + Range And DstY > SrcY - Range
YS = 0
End Select

If XS < 0 And YS < 0 Then FindAng = 7
If XS > 0 And YS > 0 Then FindAng = 3

If XS = 0 And YS > 0 Then FindAng = 2
If XS = 0 And YS < 0 Then FindAng = 6

If YS = 0 And XS > 0 Then FindAng = 4
If YS = 0 And XS < 0 Then FindAng = 0

If YS < 0 And XS > 0 Then FindAng = 5
If YS > 0 And XS < 0 Then FindAng = 1

FindAng = 7 - FindAng
End Function


Sub CheckWinner()
Dim c As Long, PCount As Long, WinnerID As Long, R, G
If Bands = False Then
For c = 1 To UBound(P())
If P(c).Act = True Then PCount = PCount + 1
Next c
If PCount <= 1 And isGameDone = True Then
For c = 1 To UBound(P())
If P(c).Act = True Then WinnerID = c
Next c
Running = False
If WinnerID <= 0 Then Exit Sub
Message.Major = P(WinnerID).Nick & " won the game!"
Message.Minor = "Press F2 for newgame!"
Playing = False
End If
End If

If Bands = True Then
For c = 1 To 4
If P(c).Act = True Then
If P(c).Band = 1 Then R = R + 1
If P(c).Band = 0 Then G = G + 1
End If
Next c
Debug.Print R
Debug.Print G
If R < 1 And G > 0 Then
'team1 won
Playing = False
Running = False
Message.Major = TeamName(1) & " won the game!"
Message.Minor = "Press F2 for newgame!"
ElseIf G < 1 And R > 0 Then
'team2 won
Message.Major = TeamName(2) & " won the game!"
Message.Minor = "Press F2 for newgame!"
Playing = False
Running = False
End If
End If
End Sub

Function isGameDone() As Boolean
isGameDone = True
End Function

Function CenterText(picture As PictureBox, text As String, offx, offy)
picture.CurrentX = picture.ScaleWidth \ 2 - picture.TextWidth(text) \ 2 + offx
picture.CurrentY = picture.ScaleHeight \ 2 - picture.TextHeight("|") \ 2 + offy
picture.Print text
End Function

Function MakeAsteroid()
Dim c As Long
For c = 0 To 10
If A(c).Act = False Then
A(c).Act = True
A(c).Rotate = Int(Rnd * 35)

With frmGame
Select Case Int(Rnd * 4)
Case 0
A(c).x = Int(Rnd * .board.ScaleWidth)
A(c).y = -15
A(c).YS = Int(Rnd * 5) + 1
A(c).XS = IIf(Int(Rnd * 1) = 0, -Int(Rnd * 5) - 1, Int(Rnd * 5) + 1)
Case 1
A(c).x = -15
A(c).y = Int(Rnd * .board.ScaleHeight)
A(c).XS = Int(Rnd * 5) + 1
A(c).YS = IIf(Int(Rnd * 1) = 0, -Int(Rnd * 5) - 1, Int(Rnd * 5) + 1)
Case 2
A(c).x = .board.ScaleWidth + 15
A(c).y = Int(Rnd * .board.ScaleWidth)
A(c).XS = -Int(Rnd * 5) + 1
A(c).YS = IIf(Int(Rnd * 1) = 0, -Int(Rnd * 5) - 1, Int(Rnd * 5) + 1)
Case 3
A(c).x = Int(Rnd * .board.ScaleWidth)
A(c).y = .board.ScaleHeight + 29
A(c).XS = IIf(Int(Rnd * 1) = 0, -Int(Rnd * 5) - 1, Int(Rnd * 5) + 1)
A(c).YS = -Int(Rnd * 5) + 1
End Select
End With
Debug.Print "Asteroid Made"
Exit For
End If
Next c
End Function

Sub Pause()
If Playing = False Then Exit Sub
Select Case Running
Case True
Running = False
Exit Sub
Case False
Running = True
Exit Sub
End Select
End Sub

Function GetIDBand(id As Long) As Long
GetIDBand = P(id).Band
End Function
