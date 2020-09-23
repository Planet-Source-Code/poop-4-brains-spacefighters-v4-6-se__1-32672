VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   15270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tekton"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   3840
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Chilada ICG Dos"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdExit 
         Caption         =   "X"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Timer docolor 
      Interval        =   10
      Left            =   3000
      Top             =   2520
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c, cstep, texti, cmode

Function MoveText()
texti = texti + 1
If texti > lstText.ListCount - 1 Then
docolor.Enabled = False
tmr.Enabled = True
End If
End Function

Private Sub cmdExit_Click()
docolor.Enabled = False
tmr.Enabled = False
Unload Me
End Sub

Private Sub docolor_Timer()
board.Cls
On Error Resume Next
Debug.Print "Rolling Credits" & Int(Rnd * 55)
If texti > lstText.ListCount - 1 Then
docolor.Enabled = False
tmr.Enabled = True
Exit Sub
End If
Dim sec() As String
c = c + cstep
If c <= 0 Then cstep = 5: MoveText
If c >= 255 Then cstep = -5
sec() = Split(lstText.list(texti), "|")
board.CurrentX = board.ScaleWidth \ 2 - board.TextWidth(sec(1)) \ 2
board.CurrentY = board.ScaleHeight \ 2 - board.TextHeight("|") \ 2
board.FontSize = Val(sec(0))

Select Case cmode
Case "BW"
board.ForeColor = RGB(c, c, c)
Case "RED"
board.ForeColor = RGB(c, 0, 0)
Case "PINK"
board.ForeColor = RGB(c * 1.4, c * 0.9, c * 0.9)
Case "GREEN"
board.ForeColor = RGB(0, c, 0)
Case "BLUE"
board.ForeColor = RGB(0, 0, c)
Case "PURPLE"
board.ForeColor = RGB(c, 0, c)
Case "YELLOW"
board.ForeColor = RGB(c, c, 0)
Case "ORANGE"
board.ForeColor = RGB(c, c * 0.7, 0)
Case "AQUA"
board.ForeColor = RGB(0, c, c)
Case Else
cmode = "BW"
End Select

board.Print sec(1)
End Sub

Private Sub Form_Load()
Dim sec() As String, sec2() As String
cstep = 5
c = 0
texti = 1
LoadList App.path + "\credits.dat", lstText
sec() = Split(lstText.list(0), "|")
board.FontName = sec(0)
board.FontBold = Val(sec(1))
cmode = sec(2)
Me.Visible = True
End Sub

Function LoadList(path As String, list As ListBox)
list.Clear
Dim n As Long

For n = 101 To 105
list.AddItem LoadResString(n)
Next n
End Function

Private Sub Form_Resize()
board.Top = 0
board.Left = 0
board.Width = Me.ScaleWidth
board.Height = Me.ScaleHeight
End Sub

Private Sub tmr_Timer()
c = c + 5
board.BackColor = RGB(c, c, c)
If c > 300 Then Unload Me: frmGame.Visible = True
End Sub

