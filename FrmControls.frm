VERSION 5.00
Begin VB.Form FrmControls 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4575
   ClientLeft      =   1725
   ClientTop       =   1305
   ClientWidth     =   2940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   196
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   480
      Left            =   450
      TabIndex        =   10
      Top             =   3900
      Width           =   2100
   End
   Begin VB.Label button 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   9
      Left            =   2310
      TabIndex        =   12
      Top             =   3000
      Width           =   390
   End
   Begin VB.Label button 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   8
      Left            =   2265
      TabIndex        =   11
      Top             =   1020
      Width           =   390
   End
   Begin VB.Label pname 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player2's Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   1
      Left            =   150
      TabIndex        =   9
      Top             =   2175
      Width           =   2670
   End
   Begin VB.Label pname 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player1's Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   0
      Left            =   165
      TabIndex        =   8
      Top             =   165
      Width           =   2670
   End
   Begin VB.Label button 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   7
      Left            =   1680
      TabIndex        =   7
      Top             =   2820
      Width           =   390
   End
   Begin VB.Label button 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   6
      Left            =   870
      TabIndex        =   6
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Label button 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   5
      Left            =   870
      TabIndex        =   5
      Top             =   2820
      Width           =   390
   End
   Begin VB.Label button 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   4
      Left            =   1275
      TabIndex        =   4
      Top             =   2610
      Width           =   390
   End
   Begin VB.Label button 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   3
      Left            =   1665
      TabIndex        =   3
      Top             =   795
      Width           =   390
   End
   Begin VB.Label button 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Ctrl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   2
      Left            =   855
      TabIndex        =   2
      Top             =   1215
      Width           =   1200
   End
   Begin VB.Label button 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   1
      Left            =   855
      TabIndex        =   1
      Top             =   795
      Width           =   390
   End
   Begin VB.Label button 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   585
      Width           =   390
   End
End
Attribute VB_Name = "FrmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width \ 2 - Me.Width \ 2
Me.Top = Screen.Height \ 2 - Me.Height \ 2
Me.Show
Pause
End Sub

Private Sub Form_Unload(Cancel As Integer)
Pause
End Sub
