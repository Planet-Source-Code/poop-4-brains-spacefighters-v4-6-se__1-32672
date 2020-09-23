VERSION 5.00
Begin VB.Form frmGraphics 
   Caption         =   "Form2"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   584
   Begin VB.PictureBox ShotSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   7
      Left            =   3600
      Picture         =   "frmGraphics.frx":0000
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   54
      Top             =   3840
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   7
      Left            =   4800
      Picture         =   "frmGraphics.frx":0B0A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   53
      Top             =   3840
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   6
      Left            =   1440
      Picture         =   "frmGraphics.frx":1614
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   52
      Top             =   1800
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   6
      Left            =   2760
      Picture         =   "frmGraphics.frx":211E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   51
      Top             =   2160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   5
      Left            =   2040
      Picture         =   "frmGraphics.frx":2C28
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   50
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   5
      Left            =   2055
      Picture         =   "frmGraphics.frx":3732
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   49
      Top             =   2280
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   4
      Left            =   2520
      Picture         =   "frmGraphics.frx":423C
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   48
      Top             =   3120
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.PictureBox ShotSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   4
      Left            =   1080
      Picture         =   "frmGraphics.frx":913E
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   47
      Top             =   3240
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.PictureBox icos 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   2
      Left            =   3120
      Picture         =   "frmGraphics.frx":E040
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   46
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox icom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   2
      Left            =   4920
      Picture         =   "frmGraphics.frx":E532
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   45
      Top             =   720
      Width           =   300
   End
   Begin VB.PictureBox pups 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   3000
      Picture         =   "frmGraphics.frx":EA24
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   44
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox pupm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   3840
      Picture         =   "frmGraphics.frx":F52E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   43
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox pupm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   3840
      Picture         =   "frmGraphics.frx":10038
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   42
      Top             =   2520
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox pups 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   3000
      Picture         =   "frmGraphics.frx":10B42
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   41
      Top             =   2520
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox icos 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   3600
      Picture         =   "frmGraphics.frx":1164C
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   40
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox icom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   3810
      Picture         =   "frmGraphics.frx":11B3E
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   39
      Top             =   570
      Width           =   300
   End
   Begin VB.PictureBox pups 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   2805
      Picture         =   "frmGraphics.frx":12030
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   38
      Top             =   1860
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox pupm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   3195
      Picture         =   "frmGraphics.frx":12B3A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   37
      Top             =   1785
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox EXPLODM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   7
      Left            =   3765
      Picture         =   "frmGraphics.frx":13644
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   36
      Top             =   2505
      Width           =   4500
   End
   Begin VB.PictureBox EXPLODS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   7
      Left            =   3720
      Picture         =   "frmGraphics.frx":19FFE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   35
      Top             =   2085
      Width           =   4500
   End
   Begin VB.PictureBox ShotMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   1080
      Picture         =   "frmGraphics.frx":209B8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   34
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   1035
      Picture         =   "frmGraphics.frx":214C2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox asts 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   5040
      Picture         =   "frmGraphics.frx":21FCC
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1620
      TabIndex        =   32
      Top             =   2760
      Width           =   24300
   End
   Begin VB.PictureBox astm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   4620
      Picture         =   "frmGraphics.frx":53D66
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1620
      TabIndex        =   31
      Top             =   1110
      Width           =   24300
   End
   Begin VB.PictureBox icom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   5040
      Picture         =   "frmGraphics.frx":85B00
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   30
      Top             =   1680
      Width           =   300
   End
   Begin VB.PictureBox icos 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   4440
      Picture         =   "frmGraphics.frx":85FF2
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   29
      Top             =   1560
      Width           =   300
   End
   Begin VB.PictureBox pupm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   6165
      Picture         =   "frmGraphics.frx":864E4
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox pups 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   5940
      Picture         =   "frmGraphics.frx":8A46E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   27
      Top             =   1020
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox EXPLODM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   6
      Left            =   2340
      Picture         =   "frmGraphics.frx":8E3F8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   26
      Top             =   2400
      Width           =   2700
   End
   Begin VB.PictureBox EXPLODS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   6
      Left            =   2115
      Picture         =   "frmGraphics.frx":92382
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   25
      Top             =   1590
      Width           =   2700
   End
   Begin VB.PictureBox EXPLODM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   5
      Left            =   2640
      Picture         =   "frmGraphics.frx":9630C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   24
      Top             =   2280
      Width           =   1800
   End
   Begin VB.PictureBox EXPLODS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   5
      Left            =   2640
      Picture         =   "frmGraphics.frx":98D7E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   1800
   End
   Begin VB.PictureBox panel 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   1500
      Picture         =   "frmGraphics.frx":9B7F0
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   22
      Top             =   690
      Width           =   1650
   End
   Begin VB.PictureBox EXPLODS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   4965
      Picture         =   "frmGraphics.frx":9EFF6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   21
      Top             =   2595
      Width           =   5400
   End
   Begin VB.PictureBox EXPLODM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   2460
      Picture         =   "frmGraphics.frx":A6EC8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   20
      Top             =   1590
      Width           =   5400
   End
   Begin VB.PictureBox EXPLODS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   1755
      Picture         =   "frmGraphics.frx":AED9A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   19
      Top             =   975
      Width           =   1815
   End
   Begin VB.PictureBox EXPLODM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   825
      Picture         =   "frmGraphics.frx":B1884
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   18
      Top             =   1935
      Width           =   1815
   End
   Begin VB.PictureBox EXPLODS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   3510
      Picture         =   "frmGraphics.frx":B436E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   17
      Top             =   1560
      Width           =   1350
   End
   Begin VB.PictureBox EXPLODM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   4920
      Picture         =   "frmGraphics.frx":B6390
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   16
      Top             =   2520
      Width           =   1350
   End
   Begin VB.PictureBox EXPLODS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   945
      Picture         =   "frmGraphics.frx":B83B2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   15
      Top             =   1785
      Width           =   2700
   End
   Begin VB.PictureBox EXPLODM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   975
      Picture         =   "frmGraphics.frx":BC33C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   14
      Top             =   2235
      Width           =   2700
   End
   Begin VB.PictureBox EXPLODM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   2190
      Picture         =   "frmGraphics.frx":C02C6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   151
      TabIndex        =   13
      Top             =   2070
      Width           =   2265
   End
   Begin VB.PictureBox EXPLODS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   2115
      Picture         =   "frmGraphics.frx":C3878
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   151
      TabIndex        =   12
      Top             =   1710
      Width           =   2265
   End
   Begin VB.PictureBox ShotSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   690
      Picture         =   "frmGraphics.frx":C6E2A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   11
      Top             =   900
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   1170
      Picture         =   "frmGraphics.frx":C7934
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   10
      Top             =   870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   1950
      Picture         =   "frmGraphics.frx":C843E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   2355
      Picture         =   "frmGraphics.frx":C8F48
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   8
      Top             =   1215
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox ShotSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   0
      Left            =   1080
      Picture         =   "frmGraphics.frx":C9A52
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.PictureBox ShotMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   0
      Left            =   1065
      Picture         =   "frmGraphics.frx":CE954
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.PictureBox ShipMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2100
      Index           =   2
      Left            =   375
      Picture         =   "frmGraphics.frx":D3856
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   5
      Top             =   495
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox ShipSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2100
      Index           =   2
      Left            =   210
      Picture         =   "frmGraphics.frx":1110F8
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   4
      Top             =   255
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox ShipMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2100
      Index           =   1
      Left            =   165
      Picture         =   "frmGraphics.frx":14E99A
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox ShipSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2100
      Index           =   1
      Left            =   165
      Picture         =   "frmGraphics.frx":18C23C
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox ShipMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2100
      Index           =   0
      Left            =   0
      Picture         =   "frmGraphics.frx":1C9ADE
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   1
      Top             =   225
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox ShipSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2100
      Index           =   0
      Left            =   60
      Picture         =   "frmGraphics.frx":207380
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   9000
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Hide
Dim c As Long
For c = 0 To 2
PMask(c) = ShipMask(c).hdc
Next c
For c = 0 To 7
SMask(c) = ShotMask(c).hdc
Next c
AMask = astm.hdc
PUPMask(0) = pupm(0).hdc
PUPMask(1) = pupm(1).hdc
End Sub

