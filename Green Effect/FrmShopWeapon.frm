VERSION 5.00
Begin VB.Form FrmShopWeapon 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   2025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Wooden Shield"
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
      Height          =   270
      Left            =   60
      TabIndex        =   16
      Top             =   3540
      Width           =   1545
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "50 Castras"
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
      Height          =   270
      Left            =   60
      TabIndex        =   15
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Metal Shield"
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
      Height          =   270
      Left            =   60
      TabIndex        =   14
      Top             =   4140
      Width           =   1425
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "100 Castras"
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
      Height          =   270
      Left            =   60
      TabIndex        =   13
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Pike"
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
      Height          =   270
      Left            =   60
      TabIndex        =   12
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "75 Castras"
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
      Height          =   270
      Left            =   60
      TabIndex        =   11
      Top             =   2100
      Width           =   1200
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Sword"
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
      Height          =   270
      Left            =   60
      TabIndex        =   10
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "150 Castras"
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
      Height          =   270
      Left            =   60
      TabIndex        =   9
      Top             =   2700
      Width           =   1320
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Bow"
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
      Height          =   270
      Left            =   60
      TabIndex        =   8
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "500 Castras"
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
      Height          =   270
      Left            =   60
      TabIndex        =   7
      Top             =   3300
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Big Stick"
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
      Height          =   270
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "10 Castras"
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
      Height          =   270
      Left            =   60
      TabIndex        =   5
      Top             =   360
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Club"
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
      Height          =   270
      Left            =   60
      TabIndex        =   4
      Top             =   660
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "20 Castras"
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
      Height          =   270
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "AXE"
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
      Height          =   270
      Left            =   60
      TabIndex        =   2
      Top             =   1260
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "40 Castras"
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
      Height          =   270
      Left            =   60
      TabIndex        =   1
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Exit"
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
      Height          =   270
      Left            =   720
      TabIndex        =   0
      Top             =   4800
      Width           =   480
   End
End
Attribute VB_Name = "FrmShopWeapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Label1_Click()
    If Castras - 10 > -1 Then
        Call BuyWeapon(1)
    Else
        Call NoSell
    End If
End Sub

Private Sub Label17_Click()
    If Castras - 50 > -1 Then
        Call BuyWeapon(7)
    Else
        Call NoSell
    End If
End Sub

Private Sub Label3_Click()
    If Castras - 20 > -1 Then
        Call BuyWeapon(2)
    Else
        Call NoSell
    End If
End Sub

Private Sub Label5_Click()
    If Castras - 40 > -1 Then
        Call BuyWeapon(3)
    Else
        Call NoSell
    End If
End Sub

Private Sub Label13_Click()
    If Castras - 75 > -1 Then
        Call BuyWeapon(4)
    Else
        Call NoSell
    End If
End Sub

Private Sub Label15_Click()
    If Castras - 100 > -1 Then
        Call BuyWeapon(8)
    Else
        Call NoSell
    End If
End Sub
Private Sub Label9_Click()
    If Castras - 500 > -1 Then
        Call BuyWeapon(6)
    Else
        Call NoSell
    End If
End Sub
Private Sub Label11_Click()
    If Castras - 150 > -1 Then
        Call BuyWeapon(5)
    Else
        Call NoSell
    End If
End Sub
Private Sub Label7_Click()
    Unload FrmShopArmour
End Sub
