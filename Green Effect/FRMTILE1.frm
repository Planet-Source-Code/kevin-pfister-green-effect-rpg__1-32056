VERSION 5.00
Begin VB.Form FrmGreenEffect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Green Effect"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "FRMTILE1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrDay 
      Interval        =   60000
      Left            =   3600
      Top             =   540
   End
   Begin VB.PictureBox PicDead 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7140
      Picture         =   "FRMTILE1.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   83
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox PicCom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   1920
      Picture         =   "FRMTILE1.frx":0865
      ScaleHeight     =   570
      ScaleWidth      =   480
      TabIndex        =   82
      Top             =   8640
      Width           =   480
   End
   Begin VB.PictureBox PicRock 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7140
      Picture         =   "FRMTILE1.frx":0D43
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   81
      Top             =   1200
      Width           =   500
   End
   Begin VB.PictureBox PicTopLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6600
      Picture         =   "FRMTILE1.frx":11D9
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   80
      Top             =   1740
      Width           =   500
   End
   Begin VB.PictureBox PicTopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6600
      Picture         =   "FRMTILE1.frx":1703
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   79
      Top             =   1200
      Width           =   500
   End
   Begin VB.PictureBox PicStopRightUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6600
      Picture         =   "FRMTILE1.frx":1C30
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   78
      Top             =   2280
      Width           =   500
   End
   Begin VB.PictureBox PicStopLeftUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   6060
      Picture         =   "FRMTILE1.frx":216C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   77
      Top             =   2280
      Width           =   495
   End
   Begin VB.PictureBox PicFenceRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   6060
      Picture         =   "FRMTILE1.frx":26AB
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   76
      Top             =   1740
      Width           =   495
   End
   Begin VB.PictureBox PicAcross 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7140
      Picture         =   "FRMTILE1.frx":2BE6
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   75
      Top             =   120
      Width           =   500
   End
   Begin VB.PictureBox PicFenceLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6600
      Picture         =   "FRMTILE1.frx":3117
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   74
      Top             =   3360
      Width           =   500
   End
   Begin VB.PictureBox PicBotRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6060
      Picture         =   "FRMTILE1.frx":3658
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   73
      Top             =   3360
      Width           =   500
   End
   Begin VB.PictureBox PicStopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   7140
      Picture         =   "FRMTILE1.frx":3B91
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   72
      Top             =   660
      Width           =   500
   End
   Begin VB.PictureBox PicStopleft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6600
      Picture         =   "FRMTILE1.frx":40B7
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   71
      Top             =   2820
      Width           =   500
   End
   Begin VB.PictureBox PicBotLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6060
      Picture         =   "FRMTILE1.frx":45E1
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   70
      Top             =   2820
      Width           =   500
   End
   Begin VB.PictureBox PicFlowers 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6600
      Picture         =   "FRMTILE1.frx":4B15
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   69
      Top             =   660
      Width           =   500
   End
   Begin VB.PictureBox PicTree 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   6600
      Picture         =   "FRMTILE1.frx":503C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   68
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PicWater 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6060
      Picture         =   "FRMTILE1.frx":558B
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   67
      Top             =   1200
      Width           =   500
   End
   Begin VB.PictureBox PicSand 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6060
      Picture         =   "FRMTILE1.frx":5A03
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   66
      Top             =   660
      Width           =   500
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   1
      Left            =   6060
      Picture         =   "FRMTILE1.frx":5F09
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   65
      Top             =   120
      Width           =   500
   End
   Begin VB.PictureBox PicMachine 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   660
      Picture         =   "FRMTILE1.frx":6435
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   64
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox PicSupportRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "FRMTILE1.frx":7107
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   58
      Top             =   8640
      Width           =   495
   End
   Begin VB.PictureBox PicSupportLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      Picture         =   "FRMTILE1.frx":7525
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   57
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicBed 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":794A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   56
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicStepsRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      Picture         =   "FRMTILE1.frx":7DB3
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   55
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicSteps 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      Picture         =   "FRMTILE1.frx":821F
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   54
      Top             =   7020
      Width           =   495
   End
   Begin VB.PictureBox PicStepsLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":85A5
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   53
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2280
      Picture         =   "FRMTILE1.frx":8A10
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   52
      Top             =   8100
      Width           =   500
   End
   Begin VB.PictureBox PicCarpetRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2820
      Picture         =   "FRMTILE1.frx":8E6E
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   51
      Top             =   8100
      Width           =   500
   End
   Begin VB.PictureBox PicBottomRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1740
      Picture         =   "FRMTILE1.frx":92C3
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   50
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetBottom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   660
      Picture         =   "FRMTILE1.frx":9783
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   49
      Top             =   8100
      Width           =   500
   End
   Begin VB.PictureBox PicCarpet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FRMTILE1.frx":9BB3
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   48
      Top             =   8100
      Width           =   500
   End
   Begin VB.PictureBox PicBottomLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1200
      Picture         =   "FRMTILE1.frx":9F10
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   47
      Top             =   8100
      Width           =   500
   End
   Begin VB.PictureBox PicCarpetTop 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      Picture         =   "FRMTILE1.frx":A3D2
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   46
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetTopLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3900
      Picture         =   "FRMTILE1.frx":A80A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   45
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetTopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4440
      Picture         =   "FRMTILE1.frx":ACC3
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   44
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicArmor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      Picture         =   "FRMTILE1.frx":B176
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   43
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicInn 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      Picture         =   "FRMTILE1.frx":B56D
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   42
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicCase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":B941
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   41
      Top             =   7020
      Width           =   495
   End
   Begin VB.PictureBox PicBCase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":BD4A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   40
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicSign 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":C1B3
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   39
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicPerson2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4380
      Picture         =   "FRMTILE1.frx":C5C5
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   38
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicPerson1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4380
      Picture         =   "FRMTILE1.frx":CA1B
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   37
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicWindow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4440
      Picture         =   "FRMTILE1.frx":CE6C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   36
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicStool 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4380
      Picture         =   "FRMTILE1.frx":D2BF
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   35
      Top             =   7020
      Width           =   495
   End
   Begin VB.PictureBox PicDoor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3900
      Picture         =   "FRMTILE1.frx":D725
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   34
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicBrick 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3900
      Picture         =   "FRMTILE1.frx":DB0C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   33
      Top             =   7020
      Width           =   495
   End
   Begin VB.PictureBox PicDirt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      Picture         =   "FRMTILE1.frx":DED9
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   32
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicDown 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   3840
      Picture         =   "FRMTILE1.frx":E306
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   31
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicDown 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   3840
      Picture         =   "FRMTILE1.frx":E770
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   30
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   3300
      Picture         =   "FRMTILE1.frx":EBD7
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   29
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   3300
      Picture         =   "FRMTILE1.frx":F02A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   28
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   2760
      Picture         =   "FRMTILE1.frx":F47D
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   27
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   2760
      Picture         =   "FRMTILE1.frx":F8B4
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   26
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   2220
      Picture         =   "FRMTILE1.frx":FD01
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   25
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   2220
      Picture         =   "FRMTILE1.frx":1014D
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   24
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicFlowers 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   3360
      Picture         =   "FRMTILE1.frx":10589
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   7020
      Width           =   500
   End
   Begin VB.Timer TmrPlayer 
      Interval        =   100
      Left            =   3600
      Top             =   60
   End
   Begin VB.PictureBox PicBotLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   120
      Picture         =   "FRMTILE1.frx":10AB0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   20
      Top             =   7020
      Width           =   500
   End
   Begin VB.PictureBox PicStopleft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   660
      Picture         =   "FRMTILE1.frx":10FE4
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   19
      Top             =   7020
      Width           =   500
   End
   Begin VB.PictureBox PicStopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   1200
      Picture         =   "FRMTILE1.frx":1150E
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   18
      Top             =   7560
      Width           =   500
   End
   Begin VB.PictureBox PicBotRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   120
      Picture         =   "FRMTILE1.frx":11A34
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   17
      Top             =   7560
      Width           =   500
   End
   Begin VB.PictureBox PicFenceLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   660
      Picture         =   "FRMTILE1.frx":11F6D
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   16
      Top             =   7560
      Width           =   500
   End
   Begin VB.PictureBox PicAcross 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   1200
      Picture         =   "FRMTILE1.frx":124AE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   7020
      Width           =   500
   End
   Begin VB.PictureBox PicFenceRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1740
      Picture         =   "FRMTILE1.frx":129DF
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   7020
      Width           =   495
   End
   Begin VB.PictureBox PicStopLeftUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1740
      Picture         =   "FRMTILE1.frx":12F1A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicChest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2820
      Picture         =   "FRMTILE1.frx":13459
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   7560
      Width           =   500
   End
   Begin VB.PictureBox PicStopRightUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   2280
      Picture         =   "FRMTILE1.frx":13951
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   7560
      Width           =   500
   End
   Begin VB.PictureBox PicTopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   2820
      Picture         =   "FRMTILE1.frx":13E8D
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   7020
      Width           =   500
   End
   Begin VB.PictureBox PicTopLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   2280
      Picture         =   "FRMTILE1.frx":143BA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   7020
      Width           =   500
   End
   Begin VB.PictureBox PicWell 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1740
      Picture         =   "FRMTILE1.frx":148E4
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicTree 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1740
      Picture         =   "FRMTILE1.frx":14E27
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicWater 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   1200
      Picture         =   "FRMTILE1.frx":15376
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   5940
      Width           =   500
   End
   Begin VB.PictureBox PicRock 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   660
      Picture         =   "FRMTILE1.frx":157EE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   6480
      Width           =   500
   End
   Begin VB.PictureBox PicSand 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   120
      Picture         =   "FRMTILE1.frx":15C84
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   6480
      Width           =   500
   End
   Begin VB.PictureBox PicPath 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   660
      Picture         =   "FRMTILE1.frx":1618A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   5940
      Width           =   500
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Index           =   0
      Left            =   120
      Picture         =   "FRMTILE1.frx":165C4
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   5940
      Width           =   500
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4260
      TabIndex        =   63
      Top             =   1380
      Width           =   1515
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Castra(s)"
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4260
      TabIndex        =   62
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Shape ProgressFore 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H00008000&
      Height          =   195
      Index           =   3
      Left            =   4260
      Top             =   1980
      Width           =   795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Health"
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4260
      TabIndex        =   61
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon"
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4260
      TabIndex        =   60
      Top             =   2220
      Width           =   1575
   End
   Begin VB.Shape ProgressBack 
      BorderColor     =   &H0000C000&
      Height          =   195
      Index           =   3
      Left            =   4260
      Top             =   1980
      Width           =   1575
   End
   Begin VB.Label LblArmour 
      BackStyle       =   0  'Transparent
      Caption         =   "Armour"
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4260
      TabIndex        =   59
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Shape ProgressFore 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H00008000&
      Height          =   195
      Index           =   2
      Left            =   4260
      Top             =   2520
      Width           =   795
   End
   Begin VB.Shape ProgressBack 
      BorderColor     =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   4260
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Shape ProgressFore 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H00008000&
      Height          =   195
      Index           =   1
      Left            =   4260
      Top             =   3060
      Width           =   795
   End
   Begin VB.Shape ProgressBack 
      BorderColor     =   &H0000C000&
      Height          =   195
      Index           =   1
      Left            =   4260
      Top             =   3060
      Width           =   1575
   End
   Begin VB.Image Imgplayer 
      Height          =   465
      Left            =   0
      Picture         =   "FRMTILE1.frx":16AF0
      Top             =   0
      Width           =   465
   End
   Begin VB.Label LblAble 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4260
      TabIndex        =   23
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label LblPos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4260
      TabIndex        =   22
      Top             =   3300
      Width           =   1575
   End
   Begin VB.Label LblText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This rpg was created by kevin pfister, in 2002. please vote for this program and have fun playing."
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   5595
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   3855
      Left            =   4140
      Top             =   180
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   1635
      Left            =   120
      Top             =   4140
      Width           =   5835
   End
   Begin VB.Label LblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Green Efffect"
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4260
      Width           =   5595
   End
End
Attribute VB_Name = "FrmGreenEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TotalMap(1 To 200) As String                'The Map array
Dim CompressedMap(1 To 200) As String           'The Map array
Dim MapGrid(1 To 10, 1 To 10) As String         'The Current Display grid
Dim FindCastra(1 To 10, 1 To 10) As Integer
Dim Inpassable(1 To 10, 1 To 10)

Dim MapPlayers(1 To 200, 1 To 200) As Integer   'The Player array
Dim PlayerName(1 To 800) As String          'The Characters Name
Dim PlayerText(1 To 800) As String          'What the characters say
Dim PlayerType(1 To 800) As Integer         'The Characters Type
Dim CharX(1 To 800) As Integer              'The X Position of the Character
Dim CharY(1 To 800) As Integer              'The Y Position of the Character
Dim HouseName(1 To 800) As String           'The name of the House (Not really Needed)
Dim HouseX(1 To 800) As Integer             'The X Position of the door
Dim HouseY(1 To 800) As Integer             'The Y position of the door
Dim MapHouse(1 To 200, 1 To 200) As Integer 'The Map X,Y of the Door
Dim BuildingX As Integer    'The Current X Building
Dim BuildingY As Integer    'The Current Y Building

Dim Color, Color2                               'The Colours
Dim r As Integer, g As Integer, b As Integer    'The (R)ed (G)reen (B)lue values
Dim R2 As Integer, G2 As Integer, B2 As Integer     'The Second (R)ed (G)reen (B)lue values

Dim ScreenX     'The Current X Screen
Dim ScreenY     'The Current Y Screen
Dim PlayerX     'The Current X Pos of the player
Dim PlayerY     'The Current Y Pos of the Player
Dim Anim        'The Current Animation of the Player
Dim InBuilding As Boolean   'Is the Player in a building?

Dim EnterX
Dim EnterY

Dim Day

Dim DNTimer

Dim CGrid(1 To 10)

Dim NiceGradient As Boolean     'Show the Nice Blending of the Tiles?


Private Sub Form_Load()
    Castras = 100
    PlayerHealth = 100
    PlayerArmour = 1
    PlayerWeapon = 1
    PlayerProgress = 1
    Day = 1
    
    Randomize           'Used in the Random Chat Sub
    PlayerX = 4         'The Default X Position of the player
    PlayerY = 4         'The Default Y Position of the player
    Imgplayer.Top = PlayerY * 375 - 240     'Set the Y position of the player image
    Imgplayer.Left = PlayerX * 375 - 135    'Set the X position of the player image
    ScreenX = 1         'The Default X Screen Position
    ScreenY = 1         'The Default Y Screen Position
    Anim = 1            'The Default Animation image
    InBuilding = False      'Set the character not to be in a building
    EnterX = PlayerX
    EnterY = PlayerY
    FrmGreenEffect.Show
    FrmGreenEffect.Visible = True
    
    NiceGradient = True 'False = Faster less graphics... True = Slower Better Grpahics
    
    Call DoTimeTravel

    'A = Wall
    'B = Door
    'C = Chest
    'D = Dirt
    'E = Grass Rocks
    'F = Flowers
    'G = Grass
    'H = Stool
    'I = Window
    'J = Person1
    'K = Person2
    'L = Book Case
    'M = Case
    'N = Carpet
    'O = Carpet Top Left
    'P = Path
    'Q = Carpet Top
    'R = Rock
    'S = Sand
    'T = Tree
    'U = Carpet Top Right
    'V = Carpet Right
    'W = Water
    'X = Carpet Bottom
    'Y = Carpet Bottom Left
    'Z = Carpet Left
    '1 = Bottom Left Fence
    '2 = Bottom Right Fence
    '3 = Stop at left Fence
    '4 = Left Fence
    '5 = Horizontal Fence
    '6 = Stop at right fence
    '7 = Right Fence
    '8 = Stop at Upper Left
    '9 = Top Left Fence
    '0 = Stop at Upper Right
    '/ = Top Right Fence
    '! = Steps Left
    '% = Steps
    '£ = Steps Right
    '$ = Bed
    '^ = Carpet Bottom Right
    '& = Support Left
    '* = Support Right
    
    Call OutofBuilding

    Open App.Path + "\GE\GEffect.Hus" For Input As #1   'Open the Door File
    'Contains the locations of the doors for entry
    Input #1, Text  'Input the number of doors
    ToNum = Val(Text)   'Set the Number
    For Extras = 1 To ToNum 'Loop to collect the door info
        For GetInfo = 1 To 3
            Input #1, Text  'input data
            If Mid(Text, 1, 6) = "##Name" Then
                CurrentHouse = CurrentHouse + 1             'New Door
                HouseName(CurrentHouse) = Mid(Text, 8)      'This is not really need, may use this later?
            ElseIf Mid(Text, 1, 6) = "##XPos" Then
                HouseX(CurrentHouse) = Val(Mid(Text, 8))    'Save X position
            ElseIf Mid(Text, 1, 6) = "##YPos" Then
                HouseY(CurrentHouse) = Val(Mid$(Text, 8))   'Save Y Position
                MapHouse(HouseX(CurrentHouse), HouseY(CurrentHouse)) = CurrentHouse
            End If
        Next
    Next
    Close
    RenderMap   'Call the main Drawing Engine
End Sub

Sub DoProgress()
    If ProgressFore(1).Width <> Int((ProgressBack(1).Width / 100) * PlayerWeapon) Then
        ProgressFore(1).Width = Int((ProgressBack(1).Width / 100) * PlayerArmour)
    End If
    If ProgressFore(2).Width <> Int((ProgressBack(2).Width / 100) * PlayerWeapon) Then
        ProgressFore(2).Width = Int((ProgressBack(2).Width / 100) * PlayerWeapon)
    End If
    If ProgressFore(3).Width <> Int((ProgressBack(3).Width / 100) * PlayerHealth) Then
        ProgressFore(3).Width = Int((ProgressBack(3).Width / 100) * PlayerHealth)
    End If
    If lblMoney.Caption <> Castras Then
        lblMoney.Caption = Castras
    End If
End Sub

Sub RenderMap()     'The Drawing Part, paints the images in the Map array to the Screen
    TmrPlayer.Enabled = False   'Stop the Player Movement Timer
    'Below converts the Chars of the Map to images...
    
    For OuterLoop = 1 To 10
        For InnerLoop = 1 To 10
            
            MapGrid(InnerLoop, OuterLoop) = Mid$(TotalMap(ScreenY * 10 - 10 + OuterLoop), ScreenX * 10 - 10 + InnerLoop, 1) 'Mid$(Map(OuterLoop), InnerLoop, 1)
            FindCastra(InnerLoop, OuterLoop) = 0    'Clear random Money
            Inpassable(InnerLoop, OuterLoop) = 0    'Clears inpassable objects
            
            If MapGrid(InnerLoop, OuterLoop) = "G" Then
                Call FrmGreenEffect.PaintPicture(PicGrass(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "W" Then
                Call FrmGreenEffect.PaintPicture(PicWater(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 2
            ElseIf MapGrid(InnerLoop, OuterLoop) = "S" Then
                Call FrmGreenEffect.PaintPicture(PicSand(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "R" Then
                Call FrmGreenEffect.PaintPicture(PicRock(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "T" Then
                Call FrmGreenEffect.PaintPicture(PicTree(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "E" Then
                Call FrmGreenEffect.PaintPicture(PicWell, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "C" Then
                Call FrmGreenEffect.PaintPicture(PicChest, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "P" Then
                Call FrmGreenEffect.PaintPicture(PicPath, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "F" Then
                Call FrmGreenEffect.PaintPicture(PicFlowers(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "D" Then
                Call FrmGreenEffect.PaintPicture(PicDirt, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "A" Then
                Call FrmGreenEffect.PaintPicture(PicBrick, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "B" Then
                Call FrmGreenEffect.PaintPicture(PicDoor, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "H" Then
                Call FrmGreenEffect.PaintPicture(PicStool, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "I" Then
                Call FrmGreenEffect.PaintPicture(PicWindow, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "1" Then
                Call FrmGreenEffect.PaintPicture(PicBotLeft(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "2" Then
                Call FrmGreenEffect.PaintPicture(PicBotRight(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "3" Then
                Call FrmGreenEffect.PaintPicture(PicStopleft(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "4" Then
                Call FrmGreenEffect.PaintPicture(PicFenceLeft(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "5" Then
                Call FrmGreenEffect.PaintPicture(PicAcross(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "6" Then
                Call FrmGreenEffect.PaintPicture(PicStopRight(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "7" Then
                Call FrmGreenEffect.PaintPicture(PicFenceRight(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "8" Then
                Call FrmGreenEffect.PaintPicture(PicStopLeftUp(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "9" Then
                Call FrmGreenEffect.PaintPicture(PicTopLeft(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "0" Then
                Call FrmGreenEffect.PaintPicture(PicStopRightUp(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "/" Then
                Call FrmGreenEffect.PaintPicture(PicTopRight(Day), InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "L" Then
                Call FrmGreenEffect.PaintPicture(PicBCase, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "M" Then
                Call FrmGreenEffect.PaintPicture(PicCase, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "N" Then
                Call FrmGreenEffect.PaintPicture(PicCarpet, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "O" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetTopLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "U" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetTopRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "V" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "^" Then
                Call FrmGreenEffect.PaintPicture(PicBottomRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "X" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetBottom, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "Y" Then
                Call FrmGreenEffect.PaintPicture(PicBottomLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "Z" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "Q" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetTop, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "!" Then
                Call FrmGreenEffect.PaintPicture(PicStepsLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "%" Then
                Call FrmGreenEffect.PaintPicture(PicSteps, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "£" Then
                Call FrmGreenEffect.PaintPicture(PicStepsRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "$" Then
                Call FrmGreenEffect.PaintPicture(PicBed, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "&" Then
                Call FrmGreenEffect.PaintPicture(PicSupportLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            ElseIf MapGrid(InnerLoop, OuterLoop) = "*" Then
                Call FrmGreenEffect.PaintPicture(PicSupportRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
                Inpassable(InnerLoop, OuterLoop) = 1
            End If
        Next
    Next
    
    FindCastra(Rnd * 9 + 1, Rnd * 9 + 1) = 1 'Random Placement of Money
    FindCastra(Rnd * 9 + 1, Rnd * 9 + 1) = 1 'Random Placement of Money
    
    If InBuilding = False And NiceGradient = True Then  'If there is need for the nice gradient then draw it
        For OutFade = 1 To 10   'Loop to blend the colours of the Images
            For Infade = 1 To 10
                If Infade - 1 > 0 Then
                    If MapGrid(Infade, OutFade) <> MapGrid(Infade - 1, OutFade) Then
                        'This is the Vertical Fade Part
                        For i = 1 To 50 Step 15
                            For ii = 1 To 375 Step 15
                                'Select the First Colour
                                If MapGrid(Infade - 1, OutFade) = "G" Then
                                    Color = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "W" Then
                                    Color = PicWater(Day).Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "S" Then
                                    Color = PicSand(Day).Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "R" Then
                                    Color = PicRock(Day).Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "T" Then
                                    Color = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "E" Then
                                    Color = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "P" Then
                                    Color = PicPath.Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "D" Then
                                    Color = PicDirt.Point(i, ii)
                                Else
                                    Color = 0
                                End If
                                'Select the Second colour
                                If MapGrid(Infade, OutFade) = "G" Then
                                    Color2 = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "W" Then
                                    Color2 = PicWater(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "S" Then
                                    Color2 = PicSand(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "R" Then
                                    Color2 = PicRock(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "T" Then
                                    Color2 = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "E" Then
                                    Color2 = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "P" Then
                                    Color2 = PicPath.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "D" Then
                                    Color2 = PicDirt.Point(i, ii)
                                Else
                                    Color2 = 0
                                End If
                                If Color <> 0 And Color2 <> 0 Then
                                    GetRgb Color, r, g, b       'Get the RGB Value
                                    GetRgb Color2, R2, G2, B2   'Get the RGB Value
                                    Percent = 100 / ((i / 50) * 100)    'The Percent that changes to create a fading effect
                                    FrmGreenEffect.PSet (375 * Infade - 135 - i, 375 * OutFade + 240 - ii), RGB(R2 - R2 / Percent + r / Percent, G2 - G2 / Percent + g / Percent, B2 - B2 / Percent + b / Percent)
                                End If
                            Next
                        Next
                    End If
                End If
                If OutFade - 1 > 0 Then
                    If MapGrid(Infade, OutFade) <> MapGrid(Infade, OutFade - 1) Then
                        'This is the Horizontal fade part
                        For i = 1 To 375 Step 15
                            For ii = 1 To 50 Step 15
                                'Get the First Colour
                                If MapGrid(Infade, OutFade - 1) = "G" Then
                                    Color = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "W" Then
                                    Color = PicWater(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "S" Then
                                    Color = PicSand(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "R" Then
                                    Color = PicRock(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "T" Then
                                    Color = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "E" Then
                                    Color = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "P" Then
                                    Color = PicPath.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "D" Then
                                    Color = PicDirt.Point(i, ii)
                                Else
                                    Color = 0
                                End If
                                'Get the econd colour
                                If MapGrid(Infade, OutFade) = "G" Then
                                    Color2 = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "W" Then
                                    Color2 = PicWater(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "S" Then
                                    Color2 = PicSand(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "R" Then
                                    Color2 = PicRock(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "T" Then
                                    Color2 = PicTree(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "E" Then
                                    Color2 = PicGrass(Day).Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "P" Then
                                    Color2 = PicPath.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "D" Then
                                    Color2 = PicDirt.Point(i, ii)
                                Else
                                    Color2 = 0
                                End If
                                If Color <> 0 And Color2 <> 0 Then
                                    GetRgb Color, r, g, b       'Get the RGB Value
                                    GetRgb Color2, R2, G2, B2   'Get the RGB Value
                                    Percent = 100 / ((ii / 50) * 100)   'The percentage use to create the fading effect
                                    FrmGreenEffect.PSet (375 * Infade + 240 - i, 375 * OutFade - 135 - ii), RGB(R2 - R2 / Percent + r / Percent, G2 - G2 / Percent + g / Percent, B2 - B2 / Percent + b / Percent)
                                End If
                            Next
                        Next
                    End If
                End If
            Next
        Next
    End If
    
    For OuterLoop = 1 To 10     'This Draws the People and Signs to the Screen, cuts out green for transparent look
        For InnerLoop = 1 To 10
            If MapPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop)) > 0 Then
                    Inpassable(InnerLoop, OuterLoop - 1) = 1
                    For InOuterLoop = 1 To 495 Step 15
                        For InInnerLoop = 1 To 495 Step 15
                            If PlayerType(MapPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 1 Then
                                Color = PicPerson1.Point(InInnerLoop, InOuterLoop)
                            ElseIf PlayerType(MapPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 2 Then
                                Color = PicPerson2.Point(InInnerLoop, InOuterLoop)
                            ElseIf PlayerType(MapPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 3 Then
                                Color = PicSign.Point(InInnerLoop, InOuterLoop)
                            End If
                            If Color <> vbGreen Then
                                FrmGreenEffect.PSet (InnerLoop * 375 - 135 + InInnerLoop, (OuterLoop - 1) * 375 - 135 + InOuterLoop), Color
                            End If
                        Next
                    Next
            End If
        Next
    Next
    TmrPlayer.Enabled = True
End Sub

Sub GetRgb(ByVal Color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
    Dim temp As Long
    temp = (Color And 255)
    red = temp And 255
    temp = Int(Color / 256)
    green = temp And 255
    temp = Int(Color / 65536)
    blue = temp And 255
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub TmrDay_Timer()
    Day = 1 - Day   'Change from day to Night
End Sub

Private Sub TmrPlayer_Timer()       'This is the player movement part
    If GetAsyncKeyState(vbKeyDown) < 0 Then
        Imgplayer.Picture = PicDown(Anim).Picture
        If InBuilding = True Then
            Call CheckBuilding
        End If
        Anim = 1 - Anim 'Change the Walking animation
        If PlayerY + 1 < 11 Then
            If Inpassable(PlayerX, PlayerY + 1) = 0 Then
                PlayerY = PlayerY + 1
                Imgplayer.Top = PlayerY * 375 - 240
                If FindCastra(PlayerX, PlayerY) = 1 Then
                    FindCastra(PlayerX, PlayerY) = 0
                    Call DoCastra(PlayerX, PlayerY)
                End If
            ElseIf Inpassable(PlayerX, PlayerY + 1) = 2 Then
                PlayerY = PlayerY + 1
                Call DoDie
            End If
        Else
            PlayerY = 1
            ScreenY = ScreenY + 1
            EnterX = PlayerX
            EnterY = PlayerY
            RenderMap   'Draw the New map
            Imgplayer.Top = PlayerY * 375 - 240
        End If
    End If
    If GetAsyncKeyState(vbKeyUp) < 0 Then
        Imgplayer.Picture = PicUp(Anim).Picture
        If InBuilding = False Then
            Call CheckForBuilding
        End If
        Anim = 1 - Anim 'Change the walking animation
        If PlayerY - 1 > 0 Then
            If Inpassable(PlayerX, PlayerY - 1) = 0 Then
                PlayerY = PlayerY - 1
                Imgplayer.Top = PlayerY * 375 - 240
                If FindCastra(PlayerX, PlayerY) = 1 Then
                    FindCastra(PlayerX, PlayerY) = 0
                    Call DoCastra(PlayerX, PlayerY)
                End If
            ElseIf Inpassable(PlayerX, PlayerY - 1) = 2 Then
                PlayerY = PlayerY - 1
                Call DoDie
            End If
        Else
            PlayerY = 10
            ScreenY = ScreenY - 1
            EnterX = PlayerX
            EnterY = PlayerY
            RenderMap   'Draw the new map
            Imgplayer.Top = PlayerY * 375 - 240
        End If
    End If
    If GetAsyncKeyState(vbKeyLeft) < 0 Then
        Imgplayer.Picture = PicLeft(Anim).Picture
        Anim = 1 - Anim 'Change the player walking animation
        If PlayerX - 1 > 0 Then
            If Inpassable(PlayerX - 1, PlayerY) = 0 Then
                PlayerX = PlayerX - 1
                Imgplayer.Left = PlayerX * 375 - 135
                If FindCastra(PlayerX, PlayerY) = 1 Then
                    FindCastra(PlayerX, PlayerY) = 0
                    Call DoCastra(PlayerX, PlayerY)
                End If
            ElseIf Inpassable(PlayerX - 1, PlayerY) = 2 Then
                PlayerX = PlayerX - 1
                Call DoDie
            End If
        Else
            PlayerX = 10
            ScreenX = ScreenX - 1
            EnterX = PlayerX
            EnterY = PlayerY
            RenderMap   'Draw the new map
            Imgplayer.Left = PlayerX * 375 - 135
        End If
    End If
    If GetAsyncKeyState(vbKeyRight) < 0 Then
        Imgplayer.Picture = PicRight(Anim).Picture
        Anim = 1 - Anim 'Change the player walking animation
        If PlayerX + 1 < 11 Then
            If Inpassable(PlayerX + 1, PlayerY) = 0 Then
                PlayerX = PlayerX + 1
                Imgplayer.Left = PlayerX * 375 - 135
                If FindCastra(PlayerX, PlayerY) = 1 Then
                    FindCastra(PlayerX, PlayerY) = 0
                    Call DoCastra(PlayerX, PlayerY)
                End If
            ElseIf Inpassable(PlayerX + 1, PlayerY) = 2 Then
                PlayerX = PlayerX + 1
                Call DoDie
            End If
        Else
            PlayerX = 1
            ScreenX = ScreenX + 1
            EnterX = PlayerX
            EnterY = PlayerY
            RenderMap   'Draw the new map
            Imgplayer.Left = PlayerX * 375 - 135
        End If
    End If
    If GetAsyncKeyState(32) < 0 Then
        If PlayerY - 1 > 0 Then
            If MapGrid(PlayerX, PlayerY - 1) = "L" Then
                Call DoBCase
            ElseIf MapGrid(PlayerX, PlayerY - 1) = "M" Then
                Call DoCase
            ElseIf MapGrid(PlayerX, PlayerY - 1) = "$" Then
                Call DoBed
            End If
        End If
    End If
    If GetAsyncKeyState(65) < 0 Then
        DoDie
    End If
    LblPos.Caption = Str$(ScreenX * 10 - 10 + PlayerX) + "," + Str$(ScreenY * 10 - 10 + PlayerY)
    DoProgress
    CheckPos
End Sub

Sub CheckPos()  'Check the player position to see if there is a message waiting
    LblAble.Caption = ""    'Clear the message label
    If MapPlayers((ScreenX * 10 - 10 + PlayerX), (ScreenY * 10 - 10 + PlayerY)) > 0 Then
        LblAble.Caption = "Message!"    'Show there is a message waiting
        If GetAsyncKeyState(32) < 0 Then    'If the Player presses space
            Call Chat((ScreenX * 10 - 10 + PlayerX), (ScreenY * 10 - 10 + PlayerY))
        End If
    End If
End Sub

Sub CheckBuilding() 'If the player is in a building, check if he is exiting
    If ScreenX = 11 And ScreenY = 11 And PlayerX = 5 And PlayerY = 10 Then
        Call OutofBuilding  'Exit the building subroutine
    End If
End Sub

Sub CheckForBuilding()  'If the player is entering a building, check for doors
    If MapHouse((ScreenX * 10 - 10 + PlayerX), (ScreenY * 10 - 10 + PlayerY)) > 0 Then
        Call BuildingMaps(MapHouse((ScreenX * 10 - 10 + PlayerX), (ScreenY * 10 - 10 + PlayerY)))
    End If
End Sub

Sub BuildingMaps(index) 'Loads and sets the new maps for the building the player enters
    TmrPlayer.Enabled = False
    BuildingX = (ScreenX * 10 - 10 + PlayerX)
    BuildingY = (ScreenY * 10 - 11 + PlayerY)
    
    FileName = App.Path + "\GE\GMap"
    
    File = FileName + Mid$(Str(index), 2)
    
    Open File + ".Map" For Input As #1
    For OuterLoop = 1 To 200
        Input #1, CompressedMap(OuterLoop)
    Next
    Close
    
    ClearData
    
    For OuterLoop = 1 To 200
        TotalMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
            
            MapPlayers(InnerLoop, OuterLoop) = 0
            
            If Mid(CompressedMap(OuterLoop), InnerLoop, 1) = "(" Then
                If Mid(CompressedMap(OuterLoop), InnerLoop + 2, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 1))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1)
                    Next
                    InnerLoop = InnerLoop + 3
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 2))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1)
                    Next
                    InnerLoop = InnerLoop + 4
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 3))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 5, 1)
                    Next
                    InnerLoop = InnerLoop + 5
                End If
            Else
                TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop, 1)
            End If
        Next
    Next
    
    Open File + ".Pls" For Input As #1
    CurrentPlayer = 0
    Input #1, Text
    ToNum = Val(Text)
    For Extras = 1 To ToNum
        For GetInfo = 1 To 5
            Input #1, Text
            If Mid(Text, 1, 6) = "##Name" Then
                CurrentPlayer = CurrentPlayer + 1
                PlayerName(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##XPos" Then
                CharX(CurrentPlayer) = Val(Mid(Text, 8))
            ElseIf Mid(Text, 1, 6) = "##YPos" Then
                CharY(CurrentPlayer) = Val(Mid$(Text, 8))
                MapPlayers(CharX(CurrentPlayer), CharY(CurrentPlayer)) = CurrentPlayer
            ElseIf Mid(Text, 1, 6) = "##Text" Then
                PlayerText(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##Type" Then
                PlayerType(CurrentPlayer) = Val(Mid(Text, 8))
            End If
        Next
    Next
    Close
    Imgplayer.Visible = False
    PlayerX = 5
    PlayerY = 11
    Imgplayer.Top = PlayerY * 375 - 240
    Imgplayer.Left = PlayerX * 375 - 135
    ScreenX = 11
    ScreenY = 11
    Anim = 1
    InBuilding = True
    RenderMap
    Imgplayer.Visible = True
    TmrPlayer.Enabled = True
End Sub

Sub DoCastra(X, Y)      'Shows the User has found a castra(Money)
    LblMessage.Caption = "Object Found..."
    LblText.Caption = "You have found a Castra"
    Castras = Castras + 1
    lblMoney.Caption = Castras
    TimeBefore = Timer
    While Timer < TimeBefore + 1
        DoEvents
    Wend
    LblMessage.Caption = ""
    LblText.Caption = ""
End Sub


Sub Chat(InnerLoop, OuterLoop)  'Shows the messages when the player talks to a person
    Dim Chat(1 To 100) As String
    If PlayerName(MapPlayers(InnerLoop, OuterLoop)) = "SignPost" Then
        LblMessage.Caption = "Reading SignPost"
    Else
        LblMessage.Caption = "Talking to " + PlayerName(MapPlayers(InnerLoop, OuterLoop))
    End If
    If PlayerText(MapPlayers(InnerLoop, OuterLoop)) = "Random" Then
        Chat(1) = "Hello"
        Chat(2) = "Go Away"
        Chat(3) = "Welcome Traveller"
        Chat(4) = "You are dressed very strange young'un"
        Chat(5) = "Welcome to Muncipium"
        Chat(6) = "Have a look around"
        Chat(7) = "Have you saved lately"
        Chat(8) = "My Name is " + PlayerName(MapPlayers(InnerLoop, OuterLoop))
        Chat(9) = "You could do with some better armour!"
        Chat(10) = "You could do with a better Weapon!"
        Chat(11) = "Try not to fall in the water, you could drown..."
        Chat(12) = "Don't fall in the holes, they are very dangerous. I have seen people die of the fall"
        
        RandomChat = Int(Rnd * 12) + 1
        LblText.Caption = Chat(RandomChat)
    Else
        LblText.Caption = PlayerText(MapPlayers(InnerLoop, OuterLoop))
    End If
    
    TimeBefore = Timer
    While Timer < TimeBefore + 1 + (Len(LblText.Caption) / 20)
        DoEvents
    Wend
    If PlayerName(MapPlayers(InnerLoop, OuterLoop)) = "ShopKeeper" Then
        TmrPlayer.Enabled = False   'Disables movement
        FrmShop.Visible = False
        FrmShop.Show
        FrmShop.Left = FrmGreenEffect.Left + Imgplayer.Left + 495
        FrmShop.Top = FrmGreenEffect.Top + Imgplayer.Top + 495
        FrmShop.Visible = True
    End If
    
    LblMessage.Caption = ""
    LblText.Caption = ""
End Sub

Sub OutofBuilding() 'Load the default map when coming out of a building
    TmrPlayer.Enabled = False
    InBuilding = False
    Open App.Path + "\GE\GEffect.Map" For Input As #1
    For OuterLoop = 1 To 200
        Input #1, CompressedMap(OuterLoop)
    Next
    Close
    
    ClearData
    
    For OuterLoop = 1 To 200
        TotalMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
        
            MapPlayers(InnerLoop, OuterLoop) = 0
            
            If Mid(CompressedMap(OuterLoop), InnerLoop, 1) = "(" Then
                If Mid(CompressedMap(OuterLoop), InnerLoop + 2, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 1))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1)
                    Next
                    InnerLoop = InnerLoop + 3
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 2))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1)
                    Next
                    InnerLoop = InnerLoop + 4
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 3))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 5, 1)
                    Next
                    InnerLoop = InnerLoop + 5
                End If
            Else
                TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop, 1)
            End If
        Next
    Next
    
    Open App.Path + "\GE\GEffect.Pls" For Input As #1
    CurrentPlayer = 0
    Input #1, Text
    ToNum = Val(Text)
    For Extras = 1 To ToNum
        For GetInfo = 1 To 5
            Input #1, Text
            If Mid(Text, 1, 6) = "##Name" Then
                CurrentPlayer = CurrentPlayer + 1
                PlayerName(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##XPos" Then
                CharX(CurrentPlayer) = Val(Mid(Text, 8))
            ElseIf Mid(Text, 1, 6) = "##YPos" Then
                CharY(CurrentPlayer) = Val(Mid$(Text, 8))
                MapPlayers(CharX(CurrentPlayer), CharY(CurrentPlayer)) = CurrentPlayer
            ElseIf Mid(Text, 1, 6) = "##Text" Then
                PlayerText(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##Type" Then
                PlayerType(CurrentPlayer) = Val(Mid(Text, 8))
            End If
        Next
    Next
    Close
    Imgplayer.Visible = False
    ScreenX = Int(BuildingX / 10)
    ScreenY = Int(BuildingY / 10)
    PlayerX = BuildingX - (ScreenX * 10)
    PlayerY = BuildingY - (ScreenY * 10)
    ScreenX = ScreenX + 1
    ScreenY = ScreenY + 1
    Imgplayer.Top = PlayerY * 375 - 240
    Imgplayer.Left = PlayerX * 375 - 135
    Anim = 1
    BuildingX = 0
    BuildingY = 0
    RenderMap
    Imgplayer.Visible = True
    TmrPlayer.Enabled = True
End Sub

Sub ClearData() 'this clears the extra character data when loading a map
    For ClearPlayer = 1 To 800
        PlayerName(ClearPlayer) = ""
        CharX(ClearPlayer) = 0
        CharY(ClearPlayer) = 0
        PlayerType(ClearPlayer) = 0
        PlayerText(ClearPlayer) = ""
    Next
End Sub

Sub DoTimeTravel()
    TmrPlayer.Enabled = False
    BuildingX = (ScreenX * 10 - 10 + PlayerX)
    BuildingY = (ScreenY * 10 - 11 + PlayerY)
    
    FileName = App.Path + "\GE\GStart"
    
    
    Open FileName + ".Map" For Input As #1
    For OuterLoop = 1 To 200
        Input #1, CompressedMap(OuterLoop)
    Next
    Close
    
    ClearData
    
    For OuterLoop = 1 To 200
        TotalMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
            
            MapPlayers(InnerLoop, OuterLoop) = 0
            
            If Mid(CompressedMap(OuterLoop), InnerLoop, 1) = "(" Then
                If Mid(CompressedMap(OuterLoop), InnerLoop + 2, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 1))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1)
                    Next
                    InnerLoop = InnerLoop + 3
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 2))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1)
                    Next
                    InnerLoop = InnerLoop + 4
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 3))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 5, 1)
                    Next
                    InnerLoop = InnerLoop + 5
                End If
            Else
                TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop, 1)
            End If
        Next
    Next
    
    Imgplayer.Visible = False
    PlayerX = 5
    PlayerY = 11
    Imgplayer.Top = PlayerY * 375 - 240
    Imgplayer.Left = PlayerX * 375 - 135
    ScreenX = 11
    ScreenY = 11
    Anim = 1
    InBuilding = True
    RenderMap
    Imgplayer.Visible = True
    TmrPlayer.Enabled = False
    
    For OuterLoop = 1 To PicMachine.Height Step 15
        For InnerLoop = 1 To PicMachine.Width Step 15
            Color = PicMachine.Point(InnerLoop, OuterLoop)
            If Color <> vbGreen Then
                FrmGreenEffect.PSet (1500 + InnerLoop, 100 + OuterLoop), Color
            End If
        Next
    Next
    For OuterLoop = 1 To PicCom.Height Step 15
        For InnerLoop = 1 To PicCom.Width Step 15
            Color = PicCom.Point(InnerLoop, OuterLoop)
            If Color <> vbGreen Then
                FrmGreenEffect.PSet (700 + InnerLoop, 500 + OuterLoop), Color
                FrmGreenEffect.PSet (700 + InnerLoop, 1200 + OuterLoop), Color
                FrmGreenEffect.PSet (700 + InnerLoop, 1900 + OuterLoop), Color
                FrmGreenEffect.PSet (700 + InnerLoop, 2600 + OuterLoop), Color
                FrmGreenEffect.PSet (3000 + InnerLoop, 500 + OuterLoop), Color
                FrmGreenEffect.PSet (3000 + InnerLoop, 1200 + OuterLoop), Color
                FrmGreenEffect.PSet (3000 + InnerLoop, 1900 + OuterLoop), Color
                FrmGreenEffect.PSet (3000 + InnerLoop, 2600 + OuterLoop), Color
            End If
        Next
    Next
    For Walk = 1 To 8
        While Timer < TimeBefore + 0.2
            DoEvents
        Wend
        Imgplayer.Picture = PicUp(Anim).Picture
        Anim = 1 - Anim 'Change the walking animation
        PlayerY = PlayerY - 1
        Imgplayer.Top = PlayerY * 375 - 240
        TimeBefore = Timer
    Next
    TmrPlayer.Enabled = True
    FrmGreenEffect.Cls
End Sub

Sub DoDie()
    Imgplayer.Visible = False
    For OuterLoop = 1 To 495 Step 15
        For InnerLoop = 1 To 495 Step 15
            Color = PicDead.Point(InnerLoop, OuterLoop)
            If Color <> vbGreen Then
                FrmGreenEffect.PSet (PlayerX * 375 - 175 + InnerLoop, PlayerY * 375 - 175 + OuterLoop), Color
            End If
        Next
    Next
    TimeBefore = Timer
    While Timer < TimeBefore + 3
        DoEvents
    Wend
    FrmGreenEffect.Cls
    Call RenderMap
    Imgplayer.Top = EnterY * 375 - 240
    Imgplayer.Left = EnterX * 375 - 240
    PlayerX = EnterX
    PlayerY = EnterY
    Imgplayer.Visible = True
End Sub
