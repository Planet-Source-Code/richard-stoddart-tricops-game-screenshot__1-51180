VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   Icon            =   "Tricops1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox bossbullmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   8760
      Picture         =   "Tricops1.frx":030A
      ScaleHeight     =   26.4
      ScaleMode       =   0  'User
      ScaleWidth      =   18
      TabIndex        =   49
      Top             =   3480
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox bossbull 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   8760
      Picture         =   "Tricops1.frx":08CC
      ScaleHeight     =   26.4
      ScaleMode       =   0  'User
      ScaleWidth      =   18
      TabIndex        =   48
      Top             =   2880
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox babymask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   7920
      Picture         =   "Tricops1.frx":0E8E
      ScaleHeight     =   1185
      ScaleWidth      =   1050
      TabIndex        =   38
      Top             =   1200
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox baby 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   9120
      Picture         =   "Tricops1.frx":503C
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox lifer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   360
      Picture         =   "Tricops1.frx":91EA
      ScaleHeight     =   420
      ScaleWidth      =   765
      TabIndex        =   36
      Top             =   1200
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox lifemask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   360
      Picture         =   "Tricops1.frx":A33C
      ScaleHeight     =   33.6
      ScaleMode       =   0  'User
      ScaleWidth      =   43.714
      TabIndex        =   35
      Top             =   1680
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox bull2MSK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   2730
      Picture         =   "Tricops1.frx":B48E
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   34
      Tag             =   "0"
      Top             =   3885
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox bull2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   1680
      Picture         =   "Tricops1.frx":B5E0
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   33
      Tag             =   "0"
      Top             =   3885
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   0
      Picture         =   "Tricops1.frx":B732
      ScaleHeight     =   1200
      ScaleWidth      =   15360
      TabIndex        =   7
      Top             =   10395
      Width           =   15390
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5040
      Top             =   3885
   End
   Begin VB.PictureBox picMissBMP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   4
      Left            =   4305
      Picture         =   "Tricops1.frx":47774
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   28
      Top             =   5985
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picMissMSK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   4
      Left            =   4305
      Picture         =   "Tricops1.frx":4A11E
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   27
      Top             =   6615
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Timer expltmr 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3990
      Top             =   3885
   End
   Begin VB.PictureBox picMissBMP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   1
      Left            =   1770
      Picture         =   "Tricops1.frx":4CAC8
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   26
      Top             =   6000
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picMissBMP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   0
      Left            =   1050
      Picture         =   "Tricops1.frx":4CD86
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picMissBMP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   2
      Left            =   2490
      Picture         =   "Tricops1.frx":4CFC9
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   24
      Top             =   6000
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picMissBMP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   3
      Left            =   3330
      Picture         =   "Tricops1.frx":4D2E2
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   23
      Top             =   6000
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picMissMSK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   0
      Left            =   1050
      Picture         =   "Tricops1.frx":4D6CF
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   22
      Top             =   6600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picMissMSK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   1
      Left            =   1650
      Picture         =   "Tricops1.frx":4D91A
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   21
      Top             =   6600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picMissMSK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   2
      Left            =   2490
      Picture         =   "Tricops1.frx":4DBE7
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   20
      Top             =   6600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picMissMSK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   3
      Left            =   3330
      Picture         =   "Tricops1.frx":50591
      ScaleHeight     =   60
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox Bomb1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   4830
      Picture         =   "Tricops1.frx":509C4
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox Bombmask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   4830
      Picture         =   "Tricops1.frx":543FE
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox plane1mask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2880
      Picture         =   "Tricops1.frx":57E38
      ScaleHeight     =   885
      ScaleWidth      =   1530
      TabIndex        =   11
      Top             =   4545
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox plane1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   1020
      Picture         =   "Tricops1.frx":5C576
      ScaleHeight     =   885
      ScaleWidth      =   1530
      TabIndex        =   10
      Top             =   4530
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox heli1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   1080
      Picture         =   "Tricops1.frx":60CB4
      ScaleHeight     =   885
      ScaleWidth      =   1530
      TabIndex        =   0
      Top             =   990
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox heli1mask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2880
      Picture         =   "Tricops1.frx":653F2
      ScaleHeight     =   885
      ScaleWidth      =   1530
      TabIndex        =   1
      Top             =   1005
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox heli2mask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2895
      Picture         =   "Tricops1.frx":69B30
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   3
      Top             =   2115
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox heli2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   1095
      Picture         =   "Tricops1.frx":6E26E
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox Bullet1mask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   2790
      Picture         =   "Tricops1.frx":729AC
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   5
      Tag             =   "0"
      Top             =   3405
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox Bullet1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   1695
      Picture         =   "Tricops1.frx":72AFE
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   4
      Tag             =   "0"
      Top             =   3420
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   1230
      Left            =   0
      ScaleHeight     =   1200
      ScaleMode       =   0  'User
      ScaleWidth      =   1021
      TabIndex        =   8
      Top             =   10395
      Width           =   15315
   End
   Begin VB.Shape bosslife 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label waypoint 
      Caption         =   "Label2"
      Height          =   375
      Index           =   8
      Left            =   9480
      TabIndex        =   47
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label waypoint 
      Caption         =   "Label2"
      Height          =   375
      Index           =   7
      Left            =   9360
      TabIndex        =   46
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label waypoint 
      Caption         =   "Label2"
      Height          =   375
      Index           =   6
      Left            =   13440
      TabIndex        =   45
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label waypoint 
      Caption         =   "Label2"
      Height          =   375
      Index           =   5
      Left            =   12960
      TabIndex        =   44
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label waypoint 
      Caption         =   "Label2"
      Height          =   375
      Index           =   4
      Left            =   13440
      TabIndex        =   43
      Top             =   9240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label waypoint 
      Caption         =   "Label2"
      Height          =   375
      Index           =   3
      Left            =   9840
      TabIndex        =   42
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label waypoint 
      Caption         =   "Label2"
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   41
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label waypoint 
      Caption         =   "Label2"
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label waypoint 
      Caption         =   "Label2"
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   39
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label fpslabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   12495
      TabIndex        =   29
      Top             =   0
      Width           =   2955
   End
   Begin VB.Label livesleft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1050
      TabIndex        =   32
      Top             =   0
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   735
      TabIndex        =   31
      Top             =   0
      Width           =   435
   End
   Begin VB.Label scorelbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   30
      Top             =   9555
      Width           =   960
   End
   Begin VB.Label starter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Index           =   2
      Left            =   6510
      TabIndex        =   18
      Top             =   4725
      Visible         =   0   'False
      Width           =   3270
   End
   Begin VB.Label starter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Set...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Index           =   1
      Left            =   6720
      TabIndex        =   17
      Top             =   4725
      Visible         =   0   'False
      Width           =   3270
   End
   Begin VB.Label starter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ready?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Index           =   0
      Left            =   6510
      TabIndex        =   16
      Top             =   4620
      Visible         =   0   'False
      Width           =   3270
   End
   Begin VB.Label planevis 
      Caption         =   "1"
      Height          =   450
      Left            =   6390
      TabIndex        =   12
      Top             =   1605
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2625
      TabIndex        =   9
      Top             =   7455
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label curpic 
      Caption         =   "1"
      Height          =   330
      Left            =   3495
      TabIndex        =   6
      Top             =   3315
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label scor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   0
      TabIndex        =   15
      Top             =   9660
      Width           =   3060
   End
   Begin VB.Shape bossbox 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   6240
      Top             =   120
      Visible         =   0   'False
      Width           =   4500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Const PIXELS = 3
Option Explicit
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Dim pbulb, bulb, bulb2, helicopter, planes, explod, j, x, y, z, r As Integer
Dim whoshot As Integer
Dim planeshoot As Integer
Dim pbulx As Integer
Dim pbuly As Integer
Dim Xdest As Integer
Dim Ydest As Integer
Dim numplanes As Integer
Dim numbullets As Integer
Dim bulx(1 To 8) As Integer
Dim buly(1 To 8) As Integer
Dim planex(1 To 10) As Integer
Dim planey(1 To 10) As Integer
Dim oldx, oldy As Integer
Dim AWidth As Integer
Dim playwha As String
Dim result As Integer
Dim rc As Integer
Dim spd As Integer
Dim dunstart As Integer
Dim ploder As Integer
Dim frame As Integer
Dim frames As Integer
Dim currentbull As Integer
Dim bossfight As Boolean
Dim GTT As Integer
Dim tempGTT As Integer
Dim babyx As Integer
Dim babyy As Integer
Dim xmov As Integer
Dim ymov As Integer
Dim curwpoint As Integer
Dim waypmove As Boolean
Dim oldwpoint As Integer
Dim bosshit As Boolean
Dim foughtyet As Boolean
Dim bossshoot As Boolean

Private Sub expltmr_Timer()
frame = frame + 1
If frame = 5 Then
   frame = 0
   expltmr.Enabled = False
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  If IsKeyDown(32) Then '(Spacebar)
Select Case KeyCode
       Case vbKeySpace
            If bulx(currentbull) = -200 Then
               s_Playsound ("Gunshot3")
               bulx(currentbull) = Xdest + 85
               buly(currentbull) = Ydest + 40
               currentbull = currentbull + 1
               If currentbull = 9 Then currentbull = 1
            End If
End Select
End Sub

Private Sub Form_Load()
Randomize Timer
Form1.Picture = LoadPicture(App.Path + "\BGpic.jpg")
getres
Me.Show
For x = 1 To 8
    bulx(x) = -200
Next x
foughtyet = False
Picture1.ScaleMode = PIXELS
Picture2.ScaleMode = PIXELS
'playwha = App.Path + "\trimusic.mid"
MsgBox "Use the UP/DOWN ARROW KEYS to move up and down, LEFT/RIGHT ARROW KEYS to speedup or slow down the chopper, and press the SPACEBAR to Fire, Have Fun!"
bossfight = False
gameloop
End Sub
Private Sub gameloop()
Call restart
Do
  Form1.Cls
  Call mainloop
    If planevis.Caption = 1 Then
     For x = 1 To numplanes
         planex(x) = planex(x) - Int(Val(Label3.Caption / 3))
         planes = BitBlt(Me.hdc, planex(x), planey(x), 110, 65, plane1mask.hdc, 0, 0, vbSrcAnd)
         planes = BitBlt(Me.hdc, planex(x), planey(x), 110, 65, plane1.hdc, 0, 0, vbSrcPaint)
         If Int(Rnd * 5) + 1 = 3 Then
          If planeshoot = 0 Then
            whoshot = Int(Rnd * numplanes) + 1
            pbulx = planex(whoshot) + 10
            pbuly = planey(whoshot) + 80
            pbulb = BitBlt(Me.hdc, pbulx, pbuly, 110, 65, bull2MSK.hdc, 0, 0, vbSrcAnd)
            pbulb = BitBlt(Me.hdc, pbulx, pbuly, 110, 65, bull2.hdc, 0, 0, vbSrcPaint)
            planeshoot = 1
          End If
         End If
         If planex(x) <= -50 Then
            planex(x) = Int(Rnd * 200) + 1000
            planey(x) = Int(Rnd * 580) + 20
            If Val(scor.Caption) > 0 Then
               scor.Caption = Val(scor.Caption) - Int(Val(Label3.Caption / 2))
            End If
            If Val(scor.Caption) <= 0 Then scor.Caption = 0
         End If
         If planex(x) <= 155 Then
            If planex(x) >= 85 Then
               If planey(x) <= Ydest + 50 Then
                  If planey(x) >= Ydest Then
                     'Helicopter has collided with a plane
                     Call crashed
                     Call restart
                  End If
               End If
            End If
         End If
         For z = 1 To UBound(bulx)
             If (bulx(z) + Bullet1.Width > planex(x)) And (bulx(z) < planex(x) + plane1.Width) And (buly(z) + Bullet1.Height > planey(x)) And (buly(z) < planey(x) + plane1.Height) Then
                oldx = planex(x)
                oldy = planey(x)
                If bulx(z) <> -200 Then
                   expltmr.Enabled = True
                   bulx(z) = -200
                   buly(z) = -200
                   planex(x) = Int(Rnd * 200) + 1000
                   planey(x) = Int(Rnd * 580) + 20
                   scor.Caption = Val(Label3.Caption * 2) + Val(scor.Caption)
                   s_Playsound ("boom2")
                End If
            End If
         Next z
     Next x
  End If
  If planeshoot = 1 Then
     pbulx = pbulx - Int(Val(Label3.Caption / 2))
     If pbulx < 0 Then planeshoot = 0
     pbulb = BitBlt(Me.hdc, pbulx, pbuly, 110, 65, bull2MSK.hdc, 0, 0, vbSrcAnd)
     pbulb = BitBlt(Me.hdc, pbulx, pbuly, 110, 65, bull2.hdc, 0, 0, vbSrcPaint)
     If pbulx < 155 Then
        If pbulx > 85 Then
           If pbuly < Ydest + 50 Then
              If pbuly > Ydest Then
                 'Helicopter has collided with a Bullett
                 Call crashed
                 Call restart
              End If
           End If
        End If
     End If
  End If
  DoEvents
  If scor.Caption >= 1000 And foughtyet = False Then
     bossfight = True
     foughtyet = True
     Call bossmode
     restart
  End If
  
Loop
End Sub
Private Sub restart()
Cls
DoEvents
spd = 0
GTT = 1
tempGTT = 1
planeshoot = 0
Label3.Caption = 10
result = mciSendString("open " & playwha & " type sequencer alias trimusic", 0&, 0, 0)
result = mciSendString("play trimusic", 0&, 0, 0)
For x = 1 To UBound(bulx)
    bulx(x) = -200
Next x
frames = 0
frame = 0
dunstart = 0
numbullets = 8
numplanes = 7
currentbull = 1
j = BitBlt(Me.hdc, 0, 0, 47, 32, lifemask.hdc, 0, 0, vbSrcAnd)
j = BitBlt(Me.hdc, 0, 0, 47, 32, lifer.hdc, 0, 0, vbSrcPaint)
For y = 0 To 2
    starter(y).Visible = True
    Call Sleep(750)
    DoEvents
    starter(y).Visible = False
Next y
If bossfight = False Then
   For x = 1 To numplanes
       planex(x) = Int(Rnd * 200) + 1000
       planey(x) = Int(Rnd * 580) + 20
   Next x
End If
Xdest = 80
Ydest = 90
curpic.Caption = 0.1

End Sub
Private Sub bossmode()
Call restart
Dim bbullx(1 To 10) As Integer
Dim bbully(1 To 10) As Integer
Dim curbossbull As Integer
Dim willshootyet As Integer
willshootyet = 0
curbossbull = 0
For x = 1 To 10
   bbullx(x) = -100
   bbully(x) = -100
Next x
bossbox.Visible = True
bosslife.Visible = True
waypmove = False
curwpoint = 0
oldwpoint = 0
babyx = waypoint(8).Left
babyy = waypoint(8).Top
bosshit = False
bossshoot = False
Do
  Cls
  Call mainloop
  For x = 1 To 10
     If bbullx(x) < 0 Then
        bbullx(x) = -100
        bbully(x) = -100
     Else
        bbullx(x) = bbullx(x) - 13
        j = BitBlt(Me.hdc, bbullx(x) - 20, bbully(x), 30, 30, bossbullmask.hdc, 0, 0, vbSrcAnd)
        j = BitBlt(Me.hdc, bbullx(x) - 20, bbully(x), 30, 30, bossbull.hdc, 0, 0, vbSrcPaint)
     End If
     If bbullx(x) < 155 Then
        If bbullx(x) > 85 Then
           If bbully(x) < Ydest + 50 Then
              If bbully(x) > Ydest Then
                 'Helicopter has collided with a Bullett
                  For y = 1 To 10
                        bbullx(y) = -100
                        bbully(y) = -100
                  Next y
                  babyx = waypoint(8).Left
                  babyy = waypoint(8).Top
                 Call crashed
                 Call restart
              End If
           End If
        End If
     End If
  Next x
  If babyx <> -100 Then
     j = BitBlt(Me.hdc, babyx, babyy, 120, 120, babymask.hdc, 0, 0, vbSrcAnd)
     j = BitBlt(Me.hdc, babyx, babyy, 120, 120, baby.hdc, 0, 0, vbSrcPaint)
  End If
  If waypmove = True And bosslife.Width > 1 Then
     If waypoint(curwpoint).Left < babyx Then babyx = babyx - 3
     If waypoint(curwpoint).Left > babyx Then babyx = babyx + 3
     If waypoint(curwpoint).Top < babyy Then babyy = babyy - 3
     If waypoint(curwpoint).Top > babyy Then babyy = babyy + 3
     If babyx > waypoint(curwpoint).Left - 5 Then
        If babyx < waypoint(curwpoint).Left + 5 Then
           If babyy > waypoint(curwpoint).Top - 5 Then
              If babyy < waypoint(curwpoint).Top + 5 Then
                 waypmove = False
              End If
           End If
        End If
     End If
  ElseIf waypmove = False And bosslife.Width > 1 Then
     'Select a random waypoint to head to
     Do
       curwpoint = Int(Rnd * 8)
     Loop Until curwpoint <> oldwpoint
     oldwpoint = curwpoint
     waypmove = True
  End If
  willshootyet = willshootyet + 1
  If willshootyet = 12 Then
     If Int(Rnd * 2) + 1 = 2 Then
        willshootyet = 0
        bossshoot = True
     End If
     willshootyet = 0
  End If
  If bossshoot = True Then
    curbossbull = curbossbull + 1
    If curbossbull = 11 Then curbossbull = 1
    If bbullx(curbossbull) = -100 Then
       bbullx(curbossbull) = babyx
       bbully(curbossbull) = babyy + (baby.Height / 2)
    End If
    bossshoot = False
  End If
  For z = 1 To UBound(bulx)
      If (bulx(z) + Bullet1.Width > babyx) And (bulx(z) < babyx + baby.Width) And (buly(z) + Bullet1.Height > babyy) And (buly(z) < babyy + baby.Height) Then
          If bulx(z) <> -200 Then
             bulx(z) = -200
             buly(z) = -200
             If bosslife.Width > 10 Then
                If bosslife.Width < 200 Then bosslife.BackColor = vbYellow
                If bosslife.Width < 100 Then bosslife.BackColor = vbRed
                bosshit = True
                Exit For
             Else
                bosslife.Width = 1
                bosshit = True
                Exit For
             End If
          End If
      End If
   Next z
   If bosshit = True And bosslife.Visible = True Then
      bosslife.Width = bosslife.Width - 10
      bosshit = False
   End If
   If bosslife.Width <= 10 And bosslife.Visible = True Then
      bosslife.Visible = False
      bossbox.Visible = False
      oldx = babyx
      oldy = babyy
      babyx = -100
      babyy = -100
      bosshit = False
      j = BitBlt(Me.hdc, babyx, babyy, 120, 120, babymask.hdc, 0, 0, vbSrcAnd)
      j = BitBlt(Me.hdc, babyx, babyy, 120, 120, baby.hdc, 0, 0, vbSrcPaint)
      expltmr.Enabled = True
      MsgBox ("Well done, you defeated the evil thingy!!, you gain double points and an extra life!!")
      scor.Caption = Val(scor.Caption) * 2
      livesleft.Caption = Val(livesleft.Caption) + 1
      bossfight = False
      Exit Sub
   End If
  DoEvents
Loop

End Sub
Private Sub Form_Unload(Cancel As Integer)
result = mciSendString("close all", 0&, 0, 0)
Unload Me
End
End Sub

Private Sub Timer1_Timer()
fpslabel = Str$(Int(frames / 3)) + " fps"
frames = 0
End Sub

Public Sub crashed()
Cls
j = BitBlt(Me.hdc, -150, Ydest, 110, 65, heli1mask.hdc, 0, 0, vbSrcAnd)
j = BitBlt(Me.hdc, -150, Ydest, 110, 65, heli1.hdc, 0, 0, vbSrcPaint)
ploder = BitBlt(Me.hdc, Xdest, Ydest, 75, 55, Bombmask.hdc, 0, 0, vbSrcAnd)
ploder = BitBlt(Me.hdc, Xdest, Ydest, 75, 55, Bomb1.hdc, 0, 0, vbSrcPaint)
j = BitBlt(Me.hdc, 0, 0, 47, 32, lifemask.hdc, 0, 0, vbSrcAnd)
j = BitBlt(Me.hdc, 0, 0, 47, 32, lifer.hdc, 0, 0, vbSrcPaint)
DoEvents
s_Playsound ("boom")
livesleft.Caption = Val(livesleft.Caption) - 1
If bossfight = False Then
   For x = 1 To numplanes
       planex(x) = Int(Rnd * 200) + 1000
       planey(x) = Int(Rnd * 580) + 20
   Next x
End If
If Val(livesleft.Caption) <= 0 Then
   MsgBox ("You  have no more choppers to ruin, GAME OVER")
   MsgBox ("You scored " & scor.Caption)
   Unload Me
End If
End Sub
Private Sub mainloop()
  curpic.Caption = Val(curpic.Caption) + 0.1
  j = BitBlt(Me.hdc, 0, 0, 47, 32, lifemask.hdc, 0, 0, vbSrcAnd)
  j = BitBlt(Me.hdc, 0, 0, 47, 32, lifer.hdc, 0, 0, vbSrcPaint)
  If curpic.Caption <= 1 And curpic.Caption >= 0 Then
     helicopter = BitBlt(Me.hdc, Xdest, Ydest, 110, 65, heli1mask.hdc, 0, 0, vbSrcAnd)
     helicopter = BitBlt(Me.hdc, Xdest, Ydest, 110, 65, heli1.hdc, 0, 0, vbSrcPaint)
  ElseIf curpic.Caption <= 2 And curpic.Caption > 1 Then
     helicopter = BitBlt(Me.hdc, Xdest, Ydest, 110, 65, heli2mask.hdc, 0, 0, vbSrcAnd)
     helicopter = BitBlt(Me.hdc, Xdest, Ydest, 110, 65, heli2.hdc, 0, 0, vbSrcPaint)
  ElseIf curpic.Caption > 2 Then
     curpic.Caption = 0.1
  End If
  For x = 1 To UBound(bulx)
     If bulx(x) <> -200 Then
        If bulx(x) > Form1.ScaleWidth Then
           bulx(x) = -200
           Exit For
        End If
        bulb = BitBlt(Me.hdc, bulx(x), buly(x), 110, 65, Bullet1mask.hdc, 0, 0, vbSrcAnd)
        bulb = BitBlt(Me.hdc, bulx(x), buly(x), 110, 65, Bullet1.hdc, 0, 0, vbSrcPaint)
        bulx(x) = bulx(x) + Int(Val(Label3.Caption) / 2)
     End If
  Next x
  spd = spd + Val(Label3.Caption)
  If spd > Picture2.ScaleWidth Then spd = 0
  If spd > (Picture2.ScaleWidth - Picture1.ScaleWidth) Then
     AWidth = Picture2.ScaleWidth - spd
     rc = BitBlt(Picture1.hdc, 0, 0, AWidth, Picture2.ScaleHeight, Picture2.hdc, spd, 0, SRCCOPY)
     rc = BitBlt(Picture1.hdc, AWidth, 0, Picture1.ScaleWidth - AWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, SRCCOPY)
  Else
     rc = BitBlt(Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, spd, 0, SRCCOPY)
  End If
  If Picture2.Visible = True Then Picture2.Visible = False
  If expltmr.Enabled = True Then
     ploder = BitBlt(Me.hdc, oldx, oldy, 75, 55, picMissMSK(frame).hdc, 0, 0, vbSrcAnd)
     ploder = BitBlt(Me.hdc, oldx, oldy, 75, 55, picMissBMP(frame).hdc, 0, 0, vbSrcPaint)
  End If
  frames = frames + 1
  If IsKeyDown(37) Then '(Left)
     If Xdest > 0 Then
        If Val(Label3.Caption) >= 10 Then
           Label3.Caption = Val(Label3.Caption) - 2
        End If
     End If
  End If
  If IsKeyDown(39) Then '(Right)
     If Xdest < 915 Then
        If Val(Label3.Caption) < 100 Then
           Label3.Caption = Val(Label3.Caption) + 2
        End If
     End If
  End If
  If IsKeyDown(38) Then '(Up)
     If Ydest > 0 Then
        Ydest = Ydest - 10
     End If
  End If
  If IsKeyDown(40) Then '(Down)
     If Ydest < 630 Then
        Ydest = Ydest + 10
     End If
  End If
  If IsKeyDown(27) Then Unload Me
  livesleft.Caption = Val(livesleft.Caption)

End Sub
