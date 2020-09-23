VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9255
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Park Avenue"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDraw 
      Interval        =   1
      Left            =   2640
      Top             =   1440
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   617
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.PictureBox closes 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         MousePointer    =   2  'Cross
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   3
         Top             =   3720
         Width           =   1455
      End
      Begin VB.PictureBox credits 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         MousePointer    =   2  'Cross
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   2
         Top             =   3720
         Width           =   1455
      End
      Begin VB.PictureBox start 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         MousePointer    =   2  'Cross
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   1
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Timer tmrUpdate 
         Interval        =   1
         Left            =   1080
         Top             =   1080
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LS = 36
Const SS = 18

Private Type MiniText
Text As String
Size As Long
End Type

Private Type Message
Text As String
Mini(1 To 5) As MiniText
TextSize As Long
End Type

Dim M As Message
Dim r1, r2, g1, g2, b1, b2

Private Sub closes_Click()
End
End Sub

Private Sub credits_Click()
tmrUpdate.Enabled = False
tmrDraw.Enabled = False
DoCredits
End Sub

Private Sub Form_Load()
Me.Visible = True

Dim X As Long

M.Text = "Spacefighters v4.6 SE"
M.TextSize = 5
M.Mini(3).Text = "By Kevin Fleet"
M.Mini(5).Text = "Copyright KevCom 2002 (R)"

For X = 1 To 5
M.Mini(X).Size = 1
Next X

r1 = 255
g1 = 255
b1 = 255
r2 = 0
g2 = 0
b2 = 0

Gradient start, dHorizontal, r2 * 0.8, g2 * 0.8, b2 * 0.8, r1 * 0.8, g1 * 0.8, b1 * 0.8
Gradient credits, dHorizontal, r2 * 0.8, g2 * 0.8, b2 * 0.8, r1 * 0.8, g1 * 0.8, b1 * 0.8
Gradient closes, dHorizontal, r2 * 0.8, g2 * 0.8, b2 * 0.8, r1 * 0.8, g1 * 0.8, b1 * 0.8

start.ForeColor = vbRed
start.FontSize = 8
start.CurrentX = start.ScaleWidth \ 2 - start.TextWidth("Start Program") \ 2
start.CurrentY = start.ScaleHeight \ 2 - start.TextHeight("|") \ 2
start.Print "Start Program"

credits.ForeColor = vbRed
credits.FontSize = 8
credits.CurrentX = credits.ScaleWidth \ 2 - credits.TextWidth("View Credits") \ 2
credits.CurrentY = credits.ScaleHeight \ 2 - credits.TextHeight("|") \ 2
credits.Print "View Credits"

closes.ForeColor = vbRed
closes.FontSize = 8
closes.CurrentX = closes.ScaleWidth \ 2 - closes.TextWidth("Exit Game") \ 2
closes.CurrentY = closes.ScaleHeight \ 2 - closes.TextHeight("|") \ 2
closes.Print "Exit Game"
End Sub

Private Sub start_Click()
Dim n As Long
tmrDraw.Enabled = False
tmrUpdate.Enabled = False

Load frmGame
frmGame.Visible = True
Unload Me
End Sub

Private Sub tmrDraw_Timer()
Dim n As Long

board.Cls

Gradient board, dHorizontal, r1, g1, b1, r2, g2, b2

Dim X As Long, OrigY
board.ForeColor = vbRed
board.FontSize = M.TextSize
board.CurrentX = board.ScaleWidth \ 2 - board.TextWidth(M.Text) \ 2
board.CurrentY = board.ScaleHeight \ 2 - (board.TextHeight("|") * 2)
board.Print M.Text

OrigY = board.ScaleHeight \ 2 - (board.TextHeight("|") * 2) + board.TextHeight("|")
For X = 1 To 5
board.FontSize = M.Mini(X).Size
board.CurrentX = board.ScaleWidth \ 2 - board.TextWidth(M.Mini(X).Text) \ 2
board.CurrentY = OrigY + ((board.TextHeight("|") + 5) * (X - 1))
board.Print M.Mini(X).Text
Next X
End Sub

Private Sub tmrUpdate_Timer()
Dim X As Long

M.TextSize = M.TextSize + 1: If M.TextSize > LS Then M.TextSize = LS

For X = 1 To 5
M.Mini(X).Size = M.Mini(X).Size + 1.6: If M.Mini(X).Size > SS Then M.Mini(X).Size = SS
Next X
End Sub
