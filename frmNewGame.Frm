VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Start A New Game"
   ClientHeight    =   5235
   ClientLeft      =   4080
   ClientTop       =   2475
   ClientWidth     =   6555
   Icon            =   "frmNewGame.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5235
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Start Game"
      Height          =   435
      Left            =   0
      TabIndex        =   11
      Top             =   4800
      Width           =   6555
   End
   Begin VB.CommandButton Load 
      Caption         =   "Load Game Setup"
      Height          =   375
      Left            =   3360
      TabIndex        =   43
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save Game Setup"
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   4440
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000080FF&
      Height          =   1170
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   6510
      TabIndex        =   37
      Top             =   3285
      Width           =   6570
      Begin MSComDlg.CommonDialog cmn 
         Left            =   4680
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Spacefighters Game (*.sgm4*)|*.sgm4*"
      End
      Begin VB.TextBox tName 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   41
         Text            =   "Team2"
         Top             =   645
         Width           =   2865
      End
      Begin VB.TextBox tName 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   40
         Text            =   "Team1"
         Top             =   120
         Width           =   2865
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Team2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   39
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Team1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.PictureBox reg3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4095
      Picture         =   "frmNewGame.frx":0442
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   35
      Top             =   5190
      Width           =   300
   End
   Begin VB.PictureBox reg2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3495
      Picture         =   "frmNewGame.frx":0934
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   34
      Top             =   5175
      Width           =   300
   End
   Begin VB.PictureBox reg1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3795
      Picture         =   "frmNewGame.frx":0E26
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   33
      Top             =   5175
      Width           =   300
   End
   Begin VB.PictureBox select1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3795
      Picture         =   "frmNewGame.frx":1318
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   32
      Top             =   4875
      Width           =   300
   End
   Begin VB.PictureBox select2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3495
      Picture         =   "frmNewGame.frx":180A
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   31
      Top             =   4890
      Width           =   300
   End
   Begin VB.PictureBox select3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4095
      Picture         =   "frmNewGame.frx":1CFC
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   30
      Top             =   4875
      Width           =   300
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H0000FF00&
      Height          =   525
      Index           =   3
      Left            =   15
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   432
      TabIndex        =   24
      Top             =   2730
      Width           =   6540
      Begin VB.TextBox txtHP 
         Height          =   285
         Index           =   3
         Left            =   6000
         TabIndex        =   47
         Text            =   "50"
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1035
         TabIndex        =   29
         Text            =   "Player4"
         Top             =   105
         Width           =   3105
      End
      Begin VB.CommandButton cmdAI 
         Caption         =   "Human"
         Height          =   255
         Index           =   3
         Left            =   45
         TabIndex        =   28
         Top             =   105
         Width           =   900
      End
      Begin VB.PictureBox Ship1 
         Height          =   345
         Index           =   3
         Left            =   4320
         Picture         =   "frmNewGame.frx":21EE
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   27
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox Ship2 
         Height          =   345
         Index           =   3
         Left            =   4680
         Picture         =   "frmNewGame.frx":26E0
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   26
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox Ship3 
         Height          =   345
         Index           =   3
         Left            =   5040
         Picture         =   "frmNewGame.frx":2BD2
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   25
         Top             =   90
         Width           =   345
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   5640
         TabIndex        =   51
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H0000FF00&
      Height          =   525
      Index           =   2
      Left            =   15
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   432
      TabIndex        =   18
      Top             =   2205
      Width           =   6540
      Begin VB.TextBox txtHP 
         Height          =   285
         Index           =   2
         Left            =   6000
         TabIndex        =   46
         Text            =   "50"
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Index           =   2
         Left            =   1035
         TabIndex        =   23
         Text            =   "Player3"
         Top             =   105
         Width           =   3105
      End
      Begin VB.CommandButton cmdAI 
         Caption         =   "Human"
         Height          =   255
         Index           =   2
         Left            =   45
         TabIndex        =   22
         Top             =   105
         Width           =   900
      End
      Begin VB.PictureBox Ship1 
         Height          =   345
         Index           =   2
         Left            =   4320
         Picture         =   "frmNewGame.frx":30C4
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   21
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox Ship2 
         Height          =   345
         Index           =   2
         Left            =   4680
         Picture         =   "frmNewGame.frx":35B6
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   20
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox Ship3 
         Height          =   345
         Index           =   2
         Left            =   5040
         Picture         =   "frmNewGame.frx":3AA8
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   19
         Top             =   90
         Width           =   345
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   50
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H0000FF00&
      Height          =   525
      Index           =   1
      Left            =   15
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   432
      TabIndex        =   12
      Top             =   1680
      Width           =   6540
      Begin VB.TextBox txtHP 
         Height          =   285
         Index           =   1
         Left            =   6000
         TabIndex        =   45
         Text            =   "50"
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   1035
         TabIndex        =   17
         Text            =   "Player2"
         Top             =   105
         Width           =   3105
      End
      Begin VB.CommandButton cmdAI 
         Caption         =   "Human"
         Height          =   255
         Index           =   1
         Left            =   45
         TabIndex        =   16
         Top             =   105
         Width           =   900
      End
      Begin VB.PictureBox Ship1 
         Height          =   345
         Index           =   1
         Left            =   4320
         Picture         =   "frmNewGame.frx":3F9A
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   15
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox Ship2 
         Height          =   345
         Index           =   1
         Left            =   4680
         Picture         =   "frmNewGame.frx":448C
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   14
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox Ship3 
         Height          =   345
         Index           =   1
         Left            =   5040
         Picture         =   "frmNewGame.frx":497E
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   13
         Top             =   90
         Width           =   345
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   49
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1125
      Left            =   4800
      ScaleHeight     =   1065
      ScaleWidth      =   1725
      TabIndex        =   6
      Top             =   15
      Width           =   1785
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "4"
         Height          =   390
         Index           =   2
         Left            =   1020
         TabIndex        =   9
         Top             =   555
         Width           =   390
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "3"
         Height          =   390
         Index           =   1
         Left            =   570
         TabIndex        =   8
         Top             =   555
         Width           =   390
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "2"
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   555
         Width           =   390
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "How Many Players?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   30
         TabIndex        =   10
         Top             =   180
         Width           =   5490
      End
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H0000FF00&
      Height          =   525
      Index           =   0
      Left            =   15
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   432
      TabIndex        =   0
      Top             =   1140
      Width           =   6540
      Begin VB.TextBox txtHP 
         Height          =   285
         Index           =   0
         Left            =   6000
         TabIndex        =   44
         Text            =   "50"
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Ship3 
         Height          =   345
         Index           =   0
         Left            =   5040
         Picture         =   "frmNewGame.frx":4E70
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   5
         Top             =   90
         Width           =   345
      End
      Begin VB.PictureBox Ship2 
         Height          =   345
         Index           =   0
         Left            =   4680
         Picture         =   "frmNewGame.frx":5362
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   4
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox Ship1 
         Height          =   345
         Index           =   0
         Left            =   4320
         Picture         =   "frmNewGame.frx":5854
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   3
         Top             =   90
         Width           =   360
      End
      Begin VB.CommandButton cmdAI 
         Caption         =   "Human"
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   2
         Top             =   105
         Width           =   900
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   1035
         TabIndex        =   1
         Text            =   "Player1"
         Top             =   105
         Width           =   3105
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   48
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Image logo 
      Height          =   870
      Left            =   3240
      Picture         =   "frmNewGame.frx":5D46
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NewGame"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   540
      Left            =   720
      TabIndex        =   36
      Top             =   240
      Width           =   2700
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c2 As Integer, Types(3) As Long

Private Sub cmdAI_Click(Index As Integer)
Select Case cmdAI(Index).Caption
Case "Human"
cmdAI(Index).Caption = "Comp"
P(Index + 1).AI = True
Exit Sub
Case "Comp"
cmdAI(Index).Caption = "Human"
P(Index + 1).AI = False
Exit Sub
End Select
End Sub

Private Sub cmdPlayer_Click(Index As Integer)
Dim c
c2 = Val(cmdPlayer(Index).Caption)

For c = 1 To 4
P(c).Act = False
picPlayer(c - 1).Visible = False
Next c

For c = 1 To c2
P(c).Act = True
picPlayer(c - 1).Visible = True
Next c
End Sub

Private Sub Command1_Click()
Dim c As Long

For c = 1 To 30
Explo(c).Act = False
Next c

For c = 0 To 20
PUP(c).Act = False
Next c

For c = 0 To 10
A(c).Act = False
Next c

ReDim S(c2 * 20) As Shot
For c = 1 To 4
P(c).Act = False
P(c).XS = 0
P(c).YS = 0
P(c).Mines = 0
P(c).Shield = 0
Next c

TeamName(1) = tName(0).text
TeamName(2) = tName(1).text

With frmGame
For c = 1 To c2
Select Case c
Case 1
MakePlayer 0, 30, txtName(c - 1).text, 0, Types(c - 1)
Case 2
MakePlayer .board.ScaleWidth - 75, .board.ScaleHeight - 100, txtName(c - 1), 0, Types(c - 1)
Case 3
MakePlayer .board.ScaleWidth - 75, 30, txtName(c - 1), 0, Types(c - 1)
Case 4
MakePlayer 0, .board.ScaleHeight - 100, txtName(c - 1), 0, Types(c - 1)
End Select
P(c).id = c
DoExplo P(c).x + 20, P(c).y + 20, 0, 0, 4, 12
P(c).Health = Val(txtHP(c - 1).text)
P(c).MAXHP = Val(txtHP(c - 1).text)
Next c
End With

Bands = False
For c = 0 To 3
If picPlayer(c).BackColor = vbRed Then Bands = True
Next c

MakePowerup 1
MakePowerup 1
MakePowerup 1
MakePowerup 0

Running = True
Playing = True
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width \ 2 - Me.Width \ 2
Me.Top = Screen.Height \ 2 - Me.Height \ 2

For c2 = 0 To 3
Types(c2) = sRed
Ship1(c2).picture = select1.picture
Next c2

cmdPlayer_Click (0)
End Sub

Private Sub Load_Click()
Dim v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant, v5 As Variant
cmn.DialogTitle = "Open Game"
cmn.ShowOpen
If Len(cmn.FileName) = 0 Then Exit Sub
Open cmn.FileName For Input As #1
Input #1, v1
c2 = v1 - 2
Call cmdPlayer_Click(c2)
For i = 1 To 4
Input #1, v1, v2, v3, v4, v5
txtName(i - 1).text = v1
P(i).AI = v2
P(i).Band = v3
Types(i - 1) = v4
txtHP(i - 1).text = v5
If P(i).AI = True Then cmdAI(i - 1).Caption = "Comp"
If P(i).AI = False Then cmdAI(i - 1).Caption = "Human"
Select Case v4
Case 0
Ship1_Click (i - 1)
Case 1
Ship2_Click (i - 1)
Case 2
Ship3_Click (i - 1)
End Select
Next i
For i = 0 To 1
Input #1, v1
tName(i).text = v1
Next i
Close #1
End Sub

Private Sub picPlayer_Click(Index As Integer)
Select Case picPlayer(Index).BackColor
Case vbGreen
picPlayer(Index).BackColor = vbRed
P(Index + 1).Band = 1
Exit Sub
Case vbRed
picPlayer(Index).BackColor = vbGreen
P(Index + 1).Band = 0
Exit Sub
End Select
End Sub

Private Sub Save_Click()
cmn.DialogTitle = "Save Game"
cmn.ShowSave
If Len(cmn.FileName) = 0 Then Exit Sub
Open cmn.FileName & ".sgm4" For Output As #1
Write #1, c2
For i = 1 To 4
Write #1, txtName(i - 1).text, P(i).AI, P(i).Band, Types(i - 1), txtHP(i - 1).text
Next i
For i = 0 To 1
Write #1, tName(i).text
Next i
Close #1
End Sub

Private Sub Ship1_Click(Index As Integer)
Types(Index) = sRed
Ship1(Index).picture = select1.picture
Ship2(Index).picture = reg2.picture
Ship3(Index).picture = reg3.picture
End Sub

Private Sub Ship2_Click(Index As Integer)
Types(Index) = sBlue
Ship1(Index).picture = reg1.picture
Ship2(Index).picture = select2.picture
Ship3(Index).picture = reg3.picture
End Sub

Private Sub Ship3_Click(Index As Integer)
Types(Index) = sGreen
Ship1(Index).picture = reg1.picture
Ship2(Index).picture = reg2.picture
Ship3(Index).picture = select3.picture
End Sub
