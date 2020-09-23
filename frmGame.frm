VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spacefighters 4"
   ClientHeight    =   7230
   ClientLeft      =   1515
   ClientTop       =   1515
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Modern"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7230
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Matisse ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   0
      Picture         =   "frmGame.frx":0E42
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9660
      Begin VB.Timer tmrFPS 
         Interval        =   1000
         Left            =   2640
         Top             =   2760
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FPS As Long

Function MainLoop()
Dim DoC As Long, c As Long, Text As String, Color As Long, c2
Do
If DoC > (10000 - Speed) And GetTickCount > 300 Then
FPS = FPS + 1
If Running = True Then
Message.Major = ""
Message.Minor = ""

DoC = 0

board.Cls

If Int(Rnd * 6) = Int(Rnd * 6) And Int(Rnd * 40) = Int(Rnd * 35) Then MakePowerup (3)
If Int(Rnd * 6) = Int(Rnd * 6) And Int(Rnd * 30) = Int(Rnd * 25) Then MakePowerup (2)
If Int(Rnd * 6) = Int(Rnd * 6) And Int(Rnd * 30) = Int(Rnd * 25) Then MakePowerup (1)
If Int(Rnd * 6) = Int(Rnd * 6) And Int(Rnd * 30) = Int(Rnd * 25) Then MakePowerup (0)
If Int(Rnd * 6) = Int(Rnd * 6) And Int(Rnd * 3) = Int(Rnd * 2) Then MakeAsteroid

With frmGraphics
For c = 1 To 4
If P(c).Act = True Then
MovePlayer P(c)

If P(c).AI = True Then
DoAI P(c)
Else
Select Case c
Case 1
DoKeys P(c), vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyControl, vbKeyDown
Case 2
DoKeys P(c), vbKeyA, vbKeyD, vbKeyW, vbKeyShift, vbKeyS
End Select
End If

BitBlt board.hdc, P(c).x, P(c).y, 75, 75, .ShipMask(P(c).Type).hdc, P(c).Rotate * 75, IIf(P(c).Boost > 0, 75, 0), SRCAND
BitBlt board.hdc, P(c).x, P(c).y, 75, 75, .ShipSprite(P(c).Type).hdc, P(c).Rotate * 75, IIf(P(c).Boost > 0, 75, 0), SRCINVERT


Select Case c
Case 1
board.CurrentX = 5
board.CurrentY = 5
BitBlt board.hdc, 0, 0, 110, 43, frmGraphics.panel.hdc, 0, 0, SRCAND
BitBlt board.hdc, 0, 0, 110, 43, frmGraphics.panel.hdc, 0, 0, SRCPAINT
Color = vbBlue
Case 2
board.CurrentX = board.ScaleWidth - board.TextWidth(Text) - 5
board.CurrentY = board.ScaleHeight - board.TextHeight("|") - 22
BitBlt board.hdc, board.ScaleWidth - 110, board.ScaleHeight - 43, 110, 43, frmGraphics.panel.hdc, 0, 0, SRCAND
BitBlt board.hdc, board.ScaleWidth - 110, board.ScaleHeight - 43, 110, 43, frmGraphics.panel.hdc, 0, 0, SRCPAINT
Color = vbRed
Case 3
board.CurrentX = board.ScaleWidth - board.TextWidth(Text) - 5
board.CurrentY = 5
BitBlt board.hdc, board.ScaleWidth - 110, 0, 110, 43, frmGraphics.panel.hdc, 0, 0, SRCAND
BitBlt board.hdc, board.ScaleWidth - 110, 0, 110, 43, frmGraphics.panel.hdc, 0, 0, SRCPAINT
Color = vbMagenta
Case 4
board.CurrentX = 5
board.CurrentY = board.ScaleHeight - board.TextHeight("|") - 22
BitBlt board.hdc, 0, board.ScaleHeight - 43, 110, 43, frmGraphics.panel.hdc, 0, 0, SRCAND
BitBlt board.hdc, 0, board.ScaleHeight - 43, 110, 43, frmGraphics.panel.hdc, 0, 0, SRCPAINT
Color = vbWhite
End Select

If P(c).Shield > 0 Then
BitBlt board.hdc, board.CurrentX, board.CurrentY + board.TextHeight("|") + 2, 20, 20, .icom(0).hdc, 0, 0, SRCAND
BitBlt board.hdc, board.CurrentX, board.CurrentY + board.TextHeight("|") + 2, 20, 20, .icos(0).hdc, 0, 0, SRCINVERT

board.Circle (P(c).x + 37, P(c).y + 37), 40, Color
End If

Select Case c
Case 1
board.CurrentX = 5
board.CurrentY = 5
Color = vbBlue
Case 2
board.CurrentX = board.ScaleWidth - board.TextWidth(Text) - 5
board.CurrentY = board.ScaleHeight - board.TextHeight("|") - 22
Color = vbRed
Case 3
board.CurrentX = board.ScaleWidth - board.TextWidth(Text) - 5
board.CurrentY = 5
Color = vbMagenta
Case 4
board.CurrentX = 5
board.CurrentY = board.ScaleHeight - board.TextHeight("|") - 22
Color = vbWhite
End Select

If P(c).Mines > 0 Then
BitBlt board.hdc, board.CurrentX + 22, board.CurrentY + board.TextHeight("|") + 2, 20, 20, .icom(1).hdc, 0, 0, SRCAND
BitBlt board.hdc, board.CurrentX + 22, board.CurrentY + board.TextHeight("|") + 2, 20, 20, .icos(1).hdc, 0, 0, SRCINVERT
End If

If P(c).Superman > 0 Then
BitBlt board.hdc, board.CurrentX + 44, board.CurrentY + board.TextHeight("|") + 2, 20, 20, .icom(2).hdc, 0, 0, SRCAND
BitBlt board.hdc, board.CurrentX + 44, board.CurrentY + board.TextHeight("|") + 2, 20, 20, .icos(2).hdc, 0, 0, SRCINVERT
End If

Text = P(c).Nick & ":             "

Select Case c
Case 1
board.CurrentX = 5
board.CurrentY = 5
Color = vbBlue
Case 2
board.CurrentX = board.ScaleWidth - board.TextWidth(Text) - 5
board.CurrentY = board.ScaleHeight - board.TextHeight("|") - 22
Color = vbRed
Case 3
board.CurrentX = board.ScaleWidth - board.TextWidth(Text) - 5
board.CurrentY = 5
Color = vbMagenta
Case 4
board.CurrentX = 5
board.CurrentY = board.ScaleHeight - board.TextHeight("|") - 22
Color = vbWhite
End Select

If Bands = True Then
Select Case P(c).Band
Case 0
board.ForeColor = vbGreen
Case 1
board.ForeColor = &H80&
End Select
Else
board.ForeColor = vbBlack
End If

Dim dx, dy, per
dx = board.CurrentX
dy = board.CurrentY

board.Print Text

dx = dx + board.TextWidth(Text) - 50

per = P(c).Health / P(c).MAXHP
per = 50 * per

board.Line (dx, dy)-(dx + 50, dy + board.TextHeight("|")), vbBlue, BF
board.Line (dx, dy)-(dx + per, dy + board.TextHeight("|")), vbGreen, BF

board.ForeColor = Color
board.FontSize = 5
board.CurrentX = P(c).x + 37 - board.TextWidth(P(c).Nick) \ 2
board.CurrentY = P(c).y + 37 - board.TextHeight("|") \ 2 - 37
board.Print P(c).Nick
board.FontSize = 10
End If
Next c

MoveShots
For c = 0 To UBound(S())
If S(c).Act = True Then
BitBlt board.hdc, S(c).x, S(c).y, 30, 30, .ShotMask(S(c).Type).hdc, S(c).Rotate * 30, 0, SRCAND
BitBlt board.hdc, S(c).x, S(c).y, 30, 30, .ShotSprite(S(c).Type).hdc, S(c).Rotate * 30, 0, SRCINVERT
If S(c).Type = 3 Then
Select Case S(c).id
Case 1
Color = vbBlue
Case 2
Color = vbRed
Case 3
Color = vbMagenta
Case 4
Color = vbWhite
End Select
board.DrawWidth = 5
board.PSet (S(c).x + 15, S(c).y + 15), Color
board.DrawWidth = 1
End If
End If
Next c

For c = 1 To 30
If Explo(c).Act = True Then
BitBlt board.hdc, Explo(c).x, Explo(c).y, 30, 30, .EXPLODM(Explo(c).Type).hdc, Explo(c).Frame * 30, 0, SRCAND
BitBlt board.hdc, Explo(c).x, Explo(c).y, 30, 30, .EXPLODS(Explo(c).Type).hdc, Explo(c).Frame * 30, 0, SRCINVERT
End If
Next c

For c = 0 To 20
If PUP(c).Act = True Then
BitBlt board.hdc, PUP(c).x, PUP(c).y, 30, 30, .pupm(PUP(c).Type).hdc, 0, 0, SRCAND
BitBlt board.hdc, PUP(c).x, PUP(c).y, 30, 30, .pups(PUP(c).Type).hdc, 0, 0, SRCINVERT
End If
Next c

For c2 = 0 To 10
If A(c2).Act = True Then
BitBlt board.hdc, A(c2).x, A(c2).y, 40, 40, frmGraphics.astm.hdc, A(c2).Rotate * 45, 0, SRCAND
BitBlt board.hdc, A(c2).x, A(c2).y, 40, 40, frmGraphics.asts.hdc, A(c2).Rotate * 45, 0, SRCINVERT
End If
Next c2

End With
CheckWinner
End If
If GetAsyncKeyState(vbKeyF2) Then Running = False: Form2.Show
If GetAsyncKeyState(vbKeyF3) Then
Speed = Val(InputBox("Please enter the new speed" & vbCrLf & "Obviously the higher the number the higher the speed", "Speed", Speed))
SaveSetting "Spacefighters", "Props", "Speed", Speed
End If
If GetAsyncKeyState(vbKeyF5) Then Load FrmControls
If GetAsyncKeyState(vbKeyF4) Then Pause


board.ForeColor = vbBlue
board.FontSize = 50
CenterText board, Message.Major, 0, 0
board.FontSize = 20
CenterText board, Message.Minor, 0, (board.TextHeight(Message.Major) * 3) + 10
board.FontSize = 10
Else
DoC = DoC + 1
End If
DoEvents
Loop
End Function

Private Sub cmdNewgame_Click()
newgame
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width \ 2 - Me.Width \ 2
Me.Top = Screen.Height \ 2 - Me.Height \ 2
If Len(GetSetting("Spacefighters", "Props", "Speed")) = 0 Then
SaveSetting "SpaceFighters", "Props", "Speed", InputBox("What speed do you want the program to run at the lower the number the slower, the higher the faster!" & vbCrLf & "1-1000", "5000", "Set Speed")
End If

Speed = Val(GetSetting("Spacefighters", "Props", "Speed"))

Randomize
Load frmGraphics
Message.Major = "Spacefighters 4"
Message.Minor = "press f2 for new game"
MainLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub tmrFPS_Timer()
Me.Caption = "Spacefighters 4 - " & FPS
FPS = 0
End Sub
