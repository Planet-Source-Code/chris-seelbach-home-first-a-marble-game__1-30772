VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home First!"
   ClientHeight    =   6105
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   7455
   ForeColor       =   &H00000000&
   Icon            =   "home1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "home1.frx":030A
   MousePointer    =   99  'Custom
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   3960
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   1920
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   1680
      Picture         =   "home1.frx":0614
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer11 
      Interval        =   2000
      Left            =   9120
      Top             =   2040
   End
   Begin VB.Timer Timer10 
      Interval        =   20
      Left            =   9120
      Top             =   1560
   End
   Begin VB.Timer Timer9 
      Interval        =   20
      Left            =   9120
      Top             =   1080
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9000
      Top             =   4080
   End
   Begin VB.Timer Timer7 
      Interval        =   1000
      Left            =   720
      Top             =   4440
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   4920
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   4920
   End
   Begin VB.CommandButton cspin 
      Caption         =   "I Spin"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   5400
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   5400
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   720
      Top             =   5880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "You Spin"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   5760
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   240
      Top             =   5880
   End
   Begin VB.Image Image7 
      Height          =   195
      Left            =   4920
      Picture         =   "home1.frx":0AC2
      Top             =   5040
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   195
      Left            =   4200
      Picture         =   "home1.frx":0FA0
      Top             =   1560
      Width           =   135
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   33
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   32
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   31
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   30
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   29
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   28
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   27
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   26
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   25
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   24
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   23
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   22
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   21
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   20
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   19
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   18
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   17
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   16
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   15
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   14
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   13
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   12
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   11
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   10
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   9
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   8
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   7
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   6
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   5
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   4
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   3
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   2
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   1
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cop 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label levh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level: Hard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Left            =   3360
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label leve 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level: Easy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Left            =   3240
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label levd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level: Difficult"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Left            =   3000
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   3600
      Picture         =   "home1.frx":147E
      Top             =   3600
      Width           =   600
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   4200
      Picture         =   "home1.frx":1BE0
      Top             =   1560
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   360
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   120
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Hint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7200
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image hintb 
      Height          =   615
      Left            =   7080
      Picture         =   "home1.frx":2342
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image whitey 
      Height          =   480
      Left            =   8400
      Picture         =   "home1.frx":2B84
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image whiter 
      Height          =   480
      Left            =   8280
      Picture         =   "home1.frx":33C6
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image whitew 
      Height          =   480
      Left            =   8160
      Picture         =   "home1.frx":3C08
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image whitep 
      Height          =   480
      Left            =   8040
      Picture         =   "home1.frx":444A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image whiteg 
      Height          =   480
      Left            =   7920
      Picture         =   "home1.frx":4C8C
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image whiteb 
      Height          =   480
      Left            =   7800
      Picture         =   "home1.frx":54CE
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image whiteh 
      Height          =   480
      Left            =   7680
      Picture         =   "home1.frx":5D10
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image mblink 
      Height          =   135
      Left            =   7320
      Top             =   5880
      Width           =   135
   End
   Begin VB.Image yblink 
      Height          =   135
      Left            =   1920
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   2280
      TabIndex        =   18
      Top             =   0
      Width           =   180
   End
   Begin VB.Image mpp 
      Height          =   480
      Left            =   9000
      Picture         =   "home1.frx":6552
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image myy 
      Enabled         =   0   'False
      Height          =   480
      Left            =   9000
      Picture         =   "home1.frx":685C
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image fpi 
      Height          =   480
      Left            =   8040
      Picture         =   "home1.frx":6B66
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image fhi 
      Height          =   480
      Left            =   8400
      Picture         =   "home1.frx":6E70
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image my 
      Enabled         =   0   'False
      Height          =   480
      Left            =   600
      Picture         =   "home1.frx":717A
      Top             =   240
      Width           =   480
   End
   Begin VB.Image mw 
      Height          =   480
      Left            =   360
      Picture         =   "home1.frx":7484
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image mr 
      Height          =   480
      Left            =   840
      Picture         =   "home1.frx":778E
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image mp 
      Height          =   480
      Left            =   8160
      Picture         =   "home1.frx":7A98
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image mg 
      Height          =   480
      Left            =   8640
      Picture         =   "home1.frx":7DA2
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image mb 
      Height          =   480
      Left            =   8280
      Picture         =   "home1.frx":80AC
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label yscore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6000
      TabIndex        =   17
      Top             =   6000
      Width           =   105
   End
   Begin VB.Label mscore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2640
      TabIndex        =   16
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4920
      TabIndex        =   15
      Top             =   6000
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1680
      TabIndex        =   14
      Top             =   1080
      Width           =   870
   End
   Begin VB.Image pm 
      Height          =   480
      Left            =   8400
      Picture         =   "home1.frx":83B6
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image rm 
      Height          =   480
      Left            =   8400
      Picture         =   "home1.frx":8BF8
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image bm 
      Height          =   480
      Left            =   8400
      Picture         =   "home1.frx":943A
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label spag 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I spin again!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2760
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Image yball 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":9C7C
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image wball 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":9F86
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image rball 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":A290
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pball 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":A59A
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image gball 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":A8A4
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image bball 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":ABAE
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   54
      Left            =   8400
      Picture         =   "home1.frx":AEB8
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   53
      Left            =   8400
      Picture         =   "home1.frx":B6FA
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   52
      Left            =   8400
      Picture         =   "home1.frx":BF3C
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   51
      Left            =   8400
      Picture         =   "home1.frx":C77E
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   50
      Left            =   8400
      Picture         =   "home1.frx":CFC0
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   49
      Left            =   8400
      Picture         =   "home1.frx":D802
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   48
      Left            =   8400
      Picture         =   "home1.frx":E044
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   47
      Left            =   8400
      Picture         =   "home1.frx":E886
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   46
      Left            =   8400
      Picture         =   "home1.frx":F0C8
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   45
      Left            =   8400
      Picture         =   "home1.frx":F90A
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   44
      Left            =   8400
      Picture         =   "home1.frx":1014C
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   43
      Left            =   8400
      Picture         =   "home1.frx":1098E
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   42
      Left            =   8400
      Picture         =   "home1.frx":111D0
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   41
      Left            =   8400
      Picture         =   "home1.frx":11A12
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   40
      Left            =   8400
      Picture         =   "home1.frx":12254
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   39
      Left            =   8400
      Picture         =   "home1.frx":12A96
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   38
      Left            =   8400
      Picture         =   "home1.frx":132D8
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image bym 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":13B1A
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image bwm 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":1435C
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image brm 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":14B9E
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image bpm 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":153E0
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image bgm 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":15C22
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image bbm 
      Height          =   480
      Left            =   720
      Picture         =   "home1.frx":16464
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label icant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I can't move!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Left            =   3120
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   37
      Left            =   8400
      Picture         =   "home1.frx":16CA6
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   36
      Left            =   8400
      Picture         =   "home1.frx":174E8
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   35
      Left            =   8400
      Picture         =   "home1.frx":17D2A
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   34
      Left            =   8400
      Picture         =   "home1.frx":1856C
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   33
      Left            =   8400
      Picture         =   "home1.frx":18DAE
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   32
      Left            =   8400
      Picture         =   "home1.frx":195F0
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label youcant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can't move!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Left            =   2760
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   3960
   End
   Begin VB.Label iwin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I Win!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   555
      Left            =   3840
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label youwin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You Win!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   555
      Left            =   3480
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   31
      Left            =   3720
      Picture         =   "home1.frx":19E32
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   30
      Left            =   4080
      Picture         =   "home1.frx":1A674
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   29
      Left            =   4440
      Picture         =   "home1.frx":1AEB6
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   28
      Left            =   4800
      Picture         =   "home1.frx":1B6F8
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image chome 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":1BF3A
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "I move:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6360
      TabIndex        =   8
      Top             =   1080
      Width           =   870
   End
   Begin VB.Image point 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":1C77C
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image mhome 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":1CBBE
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label yn 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Tag             =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.Label mn 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Tag             =   "0"
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image ehole 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":1D400
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image dragm 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":1DC42
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image home 
      Height          =   495
      Left            =   840
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Home "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5160
      TabIndex        =   4
      Top             =   3000
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Start "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4560
      TabIndex        =   3
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   27
      Left            =   5160
      Picture         =   "home1.frx":1DF4C
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   26
      Left            =   4800
      Picture         =   "home1.frx":1E78E
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   25
      Left            =   4440
      Picture         =   "home1.frx":1EFD0
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   24
      Left            =   4080
      Picture         =   "home1.frx":1F812
      Top             =   4680
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "You move:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1680
      TabIndex        =   2
      Top             =   5280
      Width           =   765
      WordWrap        =   -1  'True
   End
   Begin VB.Image m4 
      Height          =   480
      Left            =   3720
      Picture         =   "home1.frx":20054
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image m3 
      Height          =   480
      Left            =   3240
      Picture         =   "home1.frx":20896
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image m2 
      Height          =   480
      Left            =   2760
      Picture         =   "home1.frx":210D8
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image m1 
      Height          =   480
      Left            =   2280
      Picture         =   "home1.frx":2191A
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image you4 
      Height          =   480
      Left            =   5160
      Picture         =   "home1.frx":2215C
      Top             =   480
      Width           =   480
   End
   Begin VB.Image you3 
      Height          =   480
      Left            =   5640
      Picture         =   "home1.frx":2299E
      Top             =   480
      Width           =   480
   End
   Begin VB.Image you2 
      Height          =   480
      Left            =   6120
      Picture         =   "home1.frx":231E0
      Top             =   480
      Width           =   480
   End
   Begin VB.Image you1 
      Height          =   480
      Left            =   6600
      Picture         =   "home1.frx":23A22
      Top             =   480
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   23
      Left            =   3720
      Picture         =   "home1.frx":24264
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   22
      Left            =   3120
      Picture         =   "home1.frx":24AA6
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   21
      Left            =   2640
      Picture         =   "home1.frx":252E8
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   20
      Left            =   2280
      Picture         =   "home1.frx":25B2A
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   19
      Left            =   2040
      Picture         =   "home1.frx":2636C
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   18
      Left            =   1920
      Picture         =   "home1.frx":26BAE
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   17
      Left            =   2040
      Picture         =   "home1.frx":273F0
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   16
      Left            =   2280
      Picture         =   "home1.frx":27C32
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   15
      Left            =   2640
      Picture         =   "home1.frx":28474
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   14
      Left            =   3120
      Picture         =   "home1.frx":28CB6
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   13
      Left            =   3720
      Picture         =   "home1.frx":294F8
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   12
      Left            =   4440
      Picture         =   "home1.frx":29D3A
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   11
      Left            =   5160
      Picture         =   "home1.frx":2A57C
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   10
      Left            =   5760
      Picture         =   "home1.frx":2ADBE
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   9
      Left            =   6240
      Picture         =   "home1.frx":2B600
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   8
      Left            =   6600
      Picture         =   "home1.frx":2BE42
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   7
      Left            =   6840
      Picture         =   "home1.frx":2C684
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   6
      Left            =   6960
      Picture         =   "home1.frx":2CEC6
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   5
      Left            =   6840
      Picture         =   "home1.frx":2D708
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   4
      Left            =   6600
      Picture         =   "home1.frx":2DF4A
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   3
      Left            =   6240
      Picture         =   "home1.frx":2E78C
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   2
      Left            =   5760
      Picture         =   "home1.frx":2EFCE
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   1
      Left            =   5160
      Picture         =   "home1.frx":2F810
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image spot 
      Height          =   480
      Index           =   0
      Left            =   4440
      Picture         =   "home1.frx":30052
      Top             =   5280
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   6120
      TabIndex        =   0
      Top             =   0
      Width           =   180
   End
   Begin VB.Image ym 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":30894
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image wm 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":310D6
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image rdm 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":31918
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pkm 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":3215A
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image gm 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":3299C
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image blm 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":331DE
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image hole 
      Height          =   480
      Left            =   120
      Picture         =   "home1.frx":33A20
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      Height          =   6015
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   4935
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   960
      Width           =   5775
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnunulllll 
         Caption         =   "-"
      End
      Begin VB.Menu mnusaveg 
         Caption         =   "Save Game"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuopen 
         Caption         =   "Open Saved Game"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnun 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnunu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuc 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnulevel 
      Caption         =   "&Level"
      Begin VB.Menu mnueasy 
         Caption         =   "Easy"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnudiff 
         Caption         =   "Difficult"
      End
      Begin VB.Menu mnuhard 
         Caption         =   "Hard"
      End
      Begin VB.Menu mnunul 
         Caption         =   "-"
      End
      Begin VB.Menu mnuca 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuchoose 
      Caption         =   "&Choose Marble Color"
      Begin VB.Menu mnumm 
         Caption         =   "My Marbles"
         Begin VB.Menu mnublue 
            Caption         =   "Blue"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnugreen 
            Caption         =   "Green"
         End
         Begin VB.Menu mnupink 
            Caption         =   "Purple"
         End
         Begin VB.Menu mnured 
            Caption         =   "Red"
         End
         Begin VB.Menu mnuwhite 
            Caption         =   "White"
         End
         Begin VB.Menu mnuyellow 
            Caption         =   "Yellow"
         End
      End
      Begin VB.Menu mnucmarb 
         Caption         =   "Computers Marbles"
         Begin VB.Menu mnubluec 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnugreenc 
            Caption         =   "Green"
         End
         Begin VB.Menu mnupinkc 
            Caption         =   "Purple"
         End
         Begin VB.Menu mnuredc 
            Caption         =   "Red"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuwhitec 
            Caption         =   "White"
         End
         Begin VB.Menu mnuyellowc 
            Caption         =   "Yellow"
         End
      End
      Begin VB.Menu mnunull 
         Caption         =   "-"
      End
      Begin VB.Menu mnucan 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnutake 
      Caption         =   "&Take Back"
   End
   Begin VB.Menu mnupass 
      Caption         =   "&Pass"
   End
   Begin VB.Menu mnuresc 
      Caption         =   "&Reset Score"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhow 
         Caption         =   "How to play"
      End
      Begin VB.Menu mnunulll 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About Home First!"
      End
      Begin VB.Menu mnunullll 
         Caption         =   "-"
      End
      Begin VB.Menu mnucance 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim p
Dim player
Dim computer
Dim spin
Dim pickup
Dim n0
Dim n1
Dim n2
Dim n3
Dim n4
Dim n5
Dim n6
Dim n7
Dim n8
Dim n9
Dim n10
Dim n11
Dim n12
Dim n13
Dim n14
Dim n15
Dim n16
Dim n17
Dim n18
Dim n19
Dim n20
Dim n21
Dim n22
Dim n23
Dim n24
Dim n25
Dim n26
Dim six
Dim u As Integer
Dim s As Integer
Dim stuck
Dim start As Integer
Dim nomarb As Integer
Dim e As Integer
Dim cstuck
Dim f As Integer
Dim comwin As Integer
Dim blue
Dim green
Dim pink
Dim red
Dim white
Dim yellow
Dim comblue
Dim comgreen
Dim compink
Dim comred
Dim comwhite
Dim comyellow
Dim done As Integer
Dim finish As Integer
Dim motone
Dim motsix
Dim foundhit As Integer
Dim have As Integer
Dim loc1 As Integer
Dim loc2 As Integer
Dim loc3 As Integer
Dim loc4 As Integer
Dim foundhome As Integer
Dim foundhole As Integer
Dim hint As Integer
Dim noplace As Integer
Dim hit As Integer
Dim k As Integer
Dim ilpass As Integer
Dim newgame As Integer
Dim sp As Integer
Dim csg As Integer
Dim ress As Integer
Dim sgl As Integer



Private Sub Command1_Click()
Call gamel
If player = False Then Exit Sub
finish = False
Timer9.Enabled = False
Timer10.Enabled = False
mp.Visible = False
my.Visible = False
mnueasy.Enabled = False
mnudiff.Enabled = False
mnuhard.Enabled = False
mnublue.Enabled = False
mnubluec.Enabled = False
mnugreen.Enabled = False
mnugreenc.Enabled = False
mnupink.Enabled = False
mnupinkc.Enabled = False
mnured.Enabled = False
mnuredc.Enabled = False
mnuwhite.Enabled = False
mnuwhitec.Enabled = False
mnuyellow.Enabled = False
mnuyellowc.Enabled = False
icant.Visible = False
Timer3.Enabled = True
Command1.Enabled = False
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.MouseIcon = fpi.Picture
        
End Sub












Private Sub cspin_Click()
sgl = False
hit = False
leve.Visible = False
levd.Visible = False
levh.Visible = False
If mnudiff.Checked = True Then
Call comdiff
Exit Sub
ElseIf mnuhard.Checked = True Then
Call comhard
Exit Sub
Else
End If
90
youcant.Visible = False
start = False
Dim numb As Integer
Dim tnumb As Integer
Dim s As Integer
cspin.Enabled = False
Timer5.Enabled = True
Do
DoEvents
Loop Until Timer5.Enabled = False
Timer8.Enabled = True
Form1.MousePointer = 11
Do
DoEvents
Loop Until Timer8.Enabled = False
If cstuck = True Then Exit Sub
s = yn.Caption
If you1.Picture <> chome.Picture And you2.Picture <> chome.Picture And you3.Picture <> chome.Picture And you4.Picture <> chome.Picture Then
nomarb = True
GoTo 5
Else
nomarb = False
GoTo 3
End If
3
For i = 1 To 28
If spot(i - 1).Tag = yn.Caption Then
If spot(i - 1).Picture <> rm.Picture And spot(i - 1).Picture <> bm.Picture Then
spot(i - 1).Picture = rm.Picture
cop(s - 1).Picture = rm.Picture
GoTo 4
Else
GoTo 5
End If
End If
Next i
5
For i = 1 To 28
If spot(i - 1).Tag = yn.Caption Then
If spot(i - 1).Picture = bm.Picture And nomarb = False Then
spot(i - 1).Picture = rm.Picture
cop(s - 1).Picture = rm.Picture
Call evaluate
GoTo 4
Else
End If
End If
Next i
numb = 28
10
For i = 1 To 300
t = Int(Rnd * 32)
If spot(t).Tag = numb And spot(t).Picture = rm.Picture And spot(t).Tag + s < 29 Then
tnumb = spot(t).Tag + s
GoTo 16
End If
Next i
numb = numb - 1
GoTo 10
16
For j = 1 To 32
If spot(j - 1).Tag = tnumb And spot(j - 1).Picture = rm.Picture Then
numb = numb - 1
GoTo 10
ElseIf spot(j - 1).Tag = tnumb And spot(j - 1).Picture <> rm.Picture Then
spot(t).Picture = hole.Picture
cop(numb - 1).Picture = hole.Picture
Exit For
Else
End If
Next j
11
For i = 1 To 300
t = Int(Rnd * 32)
If spot(t).Tag = tnumb And spot(t).Picture <> bm.Picture And spot(t).Picture <> rm.Picture Then
spot(t).Picture = rm.Picture
cop(numb - 1 + s).Picture = rm.Picture
GoTo 17
Else
End If
If spot(t).Tag = tnumb And spot(t).Picture <> rm.Picture And spot(t).Picture <> hole.Picture Then
spot(t).Picture = rm.Picture
cop(numb - 1 + s).Picture = rm.Picture
Call evaluate
Exit For
Else
End If
Next i
17
Call win
If comwin = True Then Exit Sub
If six = True Then
Beep
spag.Visible = True
GoTo 90
Else
spag.Visible = False
End If
Timer4.Enabled = False
player = True
computer = False
mblink.Visible = True
yblink.Visible = False
Beep
Form1.MousePointer = 99
Timer2.Enabled = True
Command1.Enabled = True
Command1.SetFocus
Exit Sub
4
If you1.Picture <> ehole.Picture Then
you1.Picture = ehole.Picture
ElseIf you2.Picture <> ehole.Picture Then
you2.Picture = ehole.Picture
ElseIf you3.Picture <> ehole.Picture Then
you3.Picture = ehole.Picture
ElseIf you4.Picture <> ehole.Picture Then
you4.Picture = ehole.Picture
Else
End If
If six = True Then
Beep
spag.Visible = True
GoTo 90
Else
spag.Visible = False
End If
Timer4.Enabled = False
player = True
computer = False
mblink.Visible = True
yblink.Visible = False
Beep
Form1.MousePointer = 99
Timer2.Enabled = True
Command1.Enabled = True
Command1.SetFocus
Exit Sub
End Sub

Private Sub cspin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.MouseIcon = fpi.Picture
End Sub


Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 35
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
35
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub

Private Sub Form_Load()
Dim FName, FNum, TestString
On Error GoTo ErrorHandler
FNum = 1
FName = "saveg" & FNum & ".dat"
If Len(FName) Then
Open FName For Input As FNum
Do While Not EOF(FNum)
Input #FNum, TestString
Loop
Close
GoTo 655
End If
656
FNum = 1
TestString = 0
FName = "saveg" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
655
spot(12).Tag = 1
spot(13).Tag = 2
spot(14).Tag = 3
spot(15).Tag = 4
spot(16).Tag = 5
spot(17).Tag = 6
spot(18).Tag = 7
spot(19).Tag = 8
spot(20).Tag = 9
spot(21).Tag = 10
spot(22).Tag = 11
spot(23).Tag = 12
spot(0).Tag = 13
spot(1).Tag = 14
spot(2).Tag = 15
spot(3).Tag = 16
spot(4).Tag = 17
spot(5).Tag = 18
spot(6).Tag = 19
spot(7).Tag = 20
spot(8).Tag = 21
spot(9).Tag = 22
spot(10).Tag = 23
spot(11).Tag = 24
spot(28).Tag = 25
spot(29).Tag = 26
spot(30).Tag = 27
spot(31).Tag = 28
spot(24).Tag = 50
spot(25).Tag = 51
spot(26).Tag = 52
spot(27).Tag = 53
spot(32).Tag = 29
spot(33).Tag = 30
spot(34).Tag = 31
spot(35).Tag = 32
spot(36).Tag = 33
spot(37).Tag = 34
spot(32).Picture = pm.Picture
spot(33).Picture = pm.Picture
spot(34).Picture = pm.Picture
spot(35).Picture = pm.Picture
spot(36).Picture = pm.Picture
spot(37).Picture = pm.Picture
spot(38).Picture = pm.Picture
spot(39).Picture = pm.Picture
spot(40).Picture = pm.Picture
spot(41).Picture = pm.Picture
spot(42).Picture = pm.Picture
spot(43).Picture = pm.Picture
spot(44).Picture = pm.Picture
spot(45).Picture = pm.Picture
spot(46).Picture = pm.Picture
spot(47).Picture = pm.Picture
spot(48).Picture = pm.Picture
spot(49).Picture = pm.Picture
spot(50).Picture = pm.Picture
spot(51).Picture = pm.Picture
spot(52).Picture = pm.Picture
spot(53).Picture = pm.Picture
spot(54).Picture = pm.Picture
For i = 1 To 32
spot(i - 1).Picture = hole.Picture
Next i
m1.Picture = mhome.Picture
m2.Picture = mhome.Picture
m3.Picture = mhome.Picture
m4.Picture = mhome.Picture
you1.Picture = chome.Picture
you2.Picture = chome.Picture
you3.Picture = chome.Picture
you4.Picture = chome.Picture
player = True
finish = False
blue = True
comred = True
sgl = True
motone = 4
motsix = 1
Label6.Caption = Format(Date, "dddd, mmm d yyyy")
Label1.Caption = Format(Time, "hh:mm AM/PM")
cop(0).Tag = 1
cop(1).Tag = 2
cop(2).Tag = 3
cop(3).Tag = 4
cop(4).Tag = 5
cop(5).Tag = 6
cop(6).Tag = 7
cop(7).Tag = 8
cop(8).Tag = 9
cop(9).Tag = 10
cop(10).Tag = 11
cop(11).Tag = 12
cop(12).Tag = 13
cop(13).Tag = 14
cop(14).Tag = 15
cop(15).Tag = 16
cop(16).Tag = 17
cop(17).Tag = 18
cop(18).Tag = 19
cop(19).Tag = 20
cop(20).Tag = 21
cop(21).Tag = 22
cop(22).Tag = 23
cop(23).Tag = 24
cop(24).Tag = 25
cop(25).Tag = 26
cop(26).Tag = 27
cop(27).Tag = 28
Exit Sub
ErrorHandler:
GoTo 656
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.MouseIcon = fhi.Picture
Label7.Visible = False
End Sub


Private Sub hintb_Click()
If Timer6.Enabled = True Then Exit Sub
If hint = True Then Exit Sub
If stuck = True And youcant.Visible = True Then
MsgBox "Let me spin.", 64, "Your stuck!"
Exit Sub
End If
If icant.Visible = True Then
MsgBox "I can't win just yet, it's your turn.", 64, "I'm stuck!"
Exit Sub
End If
If player = False Then Exit Sub
If Timer3.Enabled = True Then Exit Sub
If spin = False Then
MsgBox "Spin first.", 64, "Hint"
Exit Sub
End If
hint = True
Timer12.Enabled = True
End Sub

Private Sub hintb_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 154
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
154
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub

Private Sub hintb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.MouseIcon = fpi.Picture
Label7.Visible = True
End Sub


Private Sub Image4_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 448
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
448
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Image5_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 449
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
449
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Image6_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 673
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
673
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Image7_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 674
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
674
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 70
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
70
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 64
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
64
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 67
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
67
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 66
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
66
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 68
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
68
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 106
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
106
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 107
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
107
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 108
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
108
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub levd_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 152
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
152
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub leve_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 151
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
151
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub levh_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 153
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
153
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub m1_DblClick()
If player = False Or spin = False Or m1.Picture = ehole.Picture Then Exit Sub
MsgBox "Don't click me...hold down the left mouse button to pick me up, then drag me over to where you want to move, then release the mouse button."
End Sub

Private Sub m1_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 36
If pickup = False And m1.Picture <> mhome.Picture Then
m1.Picture = mhome.Picture
Exit Sub
Else
End If
If pickup = False And m1.Picture <> ehole.Picture Then
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Source.Picture = mhome.Picture
Exit Sub
Else
End If
36
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub

Private Sub m1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If hint = True Then Exit Sub
If m1.Picture = ehole.Picture Then Exit Sub
If spin = False Then Exit Sub
If player = False Then Exit Sub
home.Picture = mhome.Picture
home.Tag = 1
m1.Picture = ehole.Picture
m1.DragIcon = dragm.Picture
m1.Drag
Image1.Tag = 100
Image3.Tag = mn.Tag
ilpass = False
End Sub


Private Sub m2_DblClick()
If player = False Or spin = False Or m2.Picture = ehole.Picture Then Exit Sub
MsgBox "Don't click me...hold down the left mouse button to pick me up, then drag me over to where you want to move, then release the mouse button."
End Sub

Private Sub m2_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 37
If pickup = False And m2.Picture <> mhome.Picture Then
m2.Picture = mhome.Picture
Exit Sub
Else
End If
If pickup = False And m2.Picture <> ehole.Picture Then
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Source.Picture = mhome.Picture
Exit Sub
Else
End If
37
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub

Private Sub m2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If hint = True Then Exit Sub
If m2.Picture = ehole.Picture Then Exit Sub
If spin = False Then Exit Sub
If player = False Then Exit Sub
home.Picture = mhome.Picture
home.Tag = 2
m2.Picture = ehole.Picture
m2.DragIcon = dragm.Picture
m2.Drag
Image1.Tag = 200
Image3.Tag = mn.Tag
ilpass = False
End Sub


Private Sub m3_DblClick()
If player = False Or spin = False Or m3.Picture = ehole.Picture Then Exit Sub
MsgBox "Don't click me...hold down the left mouse button to pick me up, then drag me over to where you want to move, then release the mouse button."
End Sub

Private Sub m3_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 38
If pickup = False And m3.Picture <> mhome.Picture Then
m3.Picture = mhome.Picture
Exit Sub
Else
End If
If pickup = False And m3.Picture <> ehole.Picture Then
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Source.Picture = mhome.Picture
Exit Sub
Else
End If
38
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub

Private Sub m3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If hint = True Then Exit Sub
If m3.Picture = ehole.Picture Then Exit Sub
If spin = False Then Exit Sub
If player = False Then Exit Sub
home.Picture = mhome.Picture
home.Tag = 3
m3.Picture = ehole.Picture
m3.DragIcon = dragm.Picture
m3.Drag
Image1.Tag = 300
Image3.Tag = mn.Tag
ilpass = False
End Sub


Private Sub m4_DblClick()
If player = False Or spin = False Or m4.Picture = ehole.Picture Then Exit Sub
MsgBox "Don't click me...hold down the left mouse button to pick me up, then drag me over to where you want to move, then release the mouse button."
End Sub

Private Sub m4_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 39
If pickup = False And m4.Picture <> mhome.Picture Then
m4.Picture = mhome.Picture
Exit Sub
Else
End If
If pickup = False And m4.Picture <> ehole.Picture Then
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Source.Picture = mhome.Picture
Exit Sub
Else
End If
39
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub

Private Sub m4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If hint = True Then Exit Sub
If m4.Picture = ehole.Picture Then Exit Sub
If spin = False Then Exit Sub
If player = False Then Exit Sub
home.Picture = mhome.Picture
home.Tag = 4
m4.Picture = ehole.Picture
m4.DragIcon = dragm.Picture
m4.Drag
Image1.Tag = 400
Image3.Tag = mn.Tag
ilpass = False
End Sub


Private Sub mblink_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 112
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
112
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub mn_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 65
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
65
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub mnuabout_Click()
MsgBox "Home First! was written for my children, all of whom love to play games. (don't we all?)  I hope you have fun playing it." _
& Chr$(13) _
& Chr$(13) _
& "By the way, if you have any problems with the game, feel free to give me a call at the house. Thanks." _
& Chr$(13) _
& Chr$(13) _
& "Chris Seelbach  Austin, Texas. (512)282-6335"
End Sub

Private Sub mnublue_Click()
If mnubluec.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
mn.ForeColor = QBColor(11)
mnublue.Checked = True
mnugreen.Checked = False
mnupink.Checked = False
mnured.Checked = False
mnuwhite.Checked = False
mnuyellow.Checked = False
dragm.Picture = bball.Picture
For i = 1 To 32
If spot(i - 1).Picture = bm.Picture Then
spot(i - 1).Picture = blm.Picture
Else
End If
Next i
bm.Picture = blm.Picture
If m1.Picture <> ehole.Picture Then
m1.Picture = bbm.Picture
Else
End If
If m2.Picture <> ehole.Picture Then
m2.Picture = bbm.Picture
Else
End If
If m3.Picture <> ehole.Picture Then
m3.Picture = bbm.Picture
Else
End If
If m4.Picture <> ehole.Picture Then
m4.Picture = bbm.Picture
Else
End If
mhome.Picture = bbm.Picture
blue = True
green = False
pink = False
red = False
white = False
yellow = False
End Sub

Private Sub mnubluec_Click()
If mnublue.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
yn.ForeColor = QBColor(11)
mnubluec.Checked = True
mnugreenc.Checked = False
mnupinkc.Checked = False
mnuredc.Checked = False
mnuwhitec.Checked = False
mnuyellowc.Checked = False
For i = 1 To 32
If spot(i - 1).Picture = rm.Picture Then
spot(i - 1).Picture = blm.Picture
Else
End If
Next i
rm.Picture = blm.Picture
If you1.Picture <> ehole.Picture Then
you1.Picture = bbm.Picture
Else
End If
If you2.Picture <> ehole.Picture Then
you2.Picture = bbm.Picture
Else
End If
If you3.Picture <> ehole.Picture Then
you3.Picture = bbm.Picture
Else
End If
If you4.Picture <> ehole.Picture Then
you4.Picture = bbm.Picture
Else
End If
chome.Picture = bbm.Picture
comblue = True
comgreen = False
compink = False
comred = False
comwhite = False
comyellow = False
End Sub


Private Sub mnudiff_Click()
mnueasy.Checked = False
mnudiff.Checked = True
mnuhard.Checked = False
End Sub

Private Sub mnueasy_Click()
mnueasy.Checked = True
mnudiff.Checked = False
mnuhard.Checked = False

End Sub

Private Sub mnugreen_Click()
If mnugreenc.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
mn.ForeColor = QBColor(10)
mnublue.Checked = False
mnugreen.Checked = True
mnupink.Checked = False
mnured.Checked = False
mnuwhite.Checked = False
mnuyellow.Checked = False
dragm.Picture = gball.Picture
For i = 1 To 32
If spot(i - 1).Picture = bm.Picture Then
spot(i - 1).Picture = gm.Picture
Else
End If
Next i
bm.Picture = gm.Picture
If m1.Picture <> ehole.Picture Then
m1.Picture = bgm.Picture
Else
End If
If m2.Picture <> ehole.Picture Then
m2.Picture = bgm.Picture
Else
End If
If m3.Picture <> ehole.Picture Then
m3.Picture = bgm.Picture
Else
End If
If m4.Picture <> ehole.Picture Then
m4.Picture = bgm.Picture
Else
End If
mhome.Picture = bgm.Picture
blue = False
green = True
pink = False
red = False
white = False
yellow = False
End Sub


Private Sub mnugreenc_Click()
If mnugreen.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
yn.ForeColor = QBColor(10)
mnubluec.Checked = False
mnugreenc.Checked = True
mnupinkc.Checked = False
mnuredc.Checked = False
mnuwhitec.Checked = False
mnuyellowc.Checked = False
For i = 1 To 32
If spot(i - 1).Picture = rm.Picture Then
spot(i - 1).Picture = gm.Picture
Else
End If
Next i
rm.Picture = gm.Picture
If you1.Picture <> ehole.Picture Then
you1.Picture = bgm.Picture
Else
End If
If you2.Picture <> ehole.Picture Then
you2.Picture = bgm.Picture
Else
End If
If you3.Picture <> ehole.Picture Then
you3.Picture = bgm.Picture
Else
End If
If you4.Picture <> ehole.Picture Then
you4.Picture = bgm.Picture
Else
End If
chome.Picture = bgm.Picture
comblue = False
comgreen = True
compink = False
comred = False
comwhite = False
comyellow = False
End Sub


Private Sub mnuhard_Click()
mnueasy.Checked = False
mnudiff.Checked = False
mnuhard.Checked = True

End Sub

Private Sub mnuhow_Click()
MsgBox "The object of this game is to move all of your marbles to home before I do. Your marbles are at the bottom of the screen, mine are at the top. You move counterclockwise around the board, just as I do. You move first. Click your spin button to get a number, then hold down the left mouse button to pick up your marble to move it. Release the mouse button to drop your marble where you want to move. You must move a marble the same number of spaces that you spin." _
& Chr$(13) _
& Chr$(13) _
& "You may move ANY one of your marbles that you wish, provided you can. You may jump over your own marbles, or mine. You can change the game level before starting a game, and change your marbles color (or mine) as well. If you land on one of my marbles it is sent back to the start, and if I land on one of yours, the same thing happens to you! You may pass your move, or take it back. If you place all four of your marbles in home before I do, you win." _
& Chr$(13) _
& Chr$(13) _
& "To get an idea of where you can move, click the question mark button and I'll show you. Good Luck!"
End Sub


Private Sub mnunew_Click()
Timer12.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer4.Enabled = False
mp.Visible = False
my.Visible = False
Timer6.Enabled = False
youwin.Visible = False
iwin.Visible = False
icant.Visible = False
youcant.Visible = False
mnueasy.Enabled = True
mnudiff.Enabled = True
mnuhard.Enabled = True
mnublue.Enabled = True
mnubluec.Enabled = True
mnugreen.Enabled = True
mnugreenc.Enabled = True
mnupink.Enabled = True
mnupinkc.Enabled = True
mnured.Enabled = True
mnuredc.Enabled = True
mnuwhite.Enabled = True
mnuwhitec.Enabled = True
mnuyellow.Enabled = True
mnuyellowc.Enabled = True
For i = 1 To 32
spot(i - 1).Picture = hole.Picture
Next i
For i = 1 To 28
cop(i - 1).Picture = hole.Picture
Next i
m1.Picture = mhome.Picture
m2.Picture = mhome.Picture
m3.Picture = mhome.Picture
m4.Picture = mhome.Picture
you1.Picture = chome.Picture
you2.Picture = chome.Picture
you3.Picture = chome.Picture
you4.Picture = chome.Picture
player = True
Timer2.Enabled = True
mblink.Visible = True
yblink.Visible = False
cspin.Enabled = False
Command1.Enabled = True
Command1.SetFocus
comwin = False
six = False
computer = False
spin = False
stuck = False
cstuck = False
pickup = False
hint = False
hit = False
finish = False
mn.Caption = 0
yn.Caption = 0
leve.Visible = False
levd.Visible = False
levh.Visible = False
newgame = True
End Sub

Private Sub mnuopen_Click()
Dim FileName, TextData, FNum
FNum = 1
FileName = "saveg" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Do While Not EOF(FNum)
Input #FNum, TextData
Loop
sp = TextData
Close
End If
If sp = 0 Then
MsgBox "You don't have a saved game.", 64
Exit Sub
Else
End If
FNum = 1
FileName = "levlg" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Do While Not EOF(FNum)
Input #FNum, TextData
Loop
sp = TextData
Close
End If
If sp = 1 Then
mnueasy.Checked = True
mnudiff.Checked = False
mnuhard.Checked = False
ElseIf sp = 2 Then
mnudiff.Checked = True
mnuhard.Checked = False
mnueasy.Checked = False
ElseIf sp = 3 Then
mnuhard.Checked = True
mnudiff.Checked = False
mnueasy.Checked = False
Else
End If
FNum = 1
FileName = "scorm" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Do While Not EOF(FNum)
Input #FNum, TextData
Loop
sp = TextData
yscore.Caption = sp
Close
End If
FNum = 1
FileName = "scorc" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Do While Not EOF(FNum)
Input #FNum, TextData
Loop
sp = TextData
mscore.Caption = sp
Close
End If
mnueasy.Enabled = False
mnudiff.Enabled = False
mnuhard.Enabled = False
mnublue.Enabled = False
mnubluec.Enabled = False
mnugreen.Enabled = False
mnugreenc.Enabled = False
mnupink.Enabled = False
mnupinkc.Enabled = False
mnured.Enabled = False
mnuredc.Enabled = False
mnuwhite.Enabled = False
mnuwhitec.Enabled = False
mnuyellow.Enabled = False
mnuyellowc.Enabled = False
Timer12.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer4.Enabled = False
mp.Visible = False
my.Visible = False
Timer6.Enabled = False
youwin.Visible = False
iwin.Visible = False
icant.Visible = False
youcant.Visible = False
player = True
Timer2.Enabled = True
mblink.Visible = True
yblink.Visible = False
cspin.Enabled = False
Command1.Enabled = True
Command1.SetFocus
comwin = False
six = False
computer = False
spin = False
stuck = False
cstuck = False
pickup = False
hint = False
hit = False
finish = False
mn.Caption = 0
yn.Caption = 0
leve.Visible = False
levd.Visible = False
levh.Visible = False
newgame = True
On Error GoTo ErrorHandler
For i = 1 To 32
spot(i - 1).Picture = hole.Picture
Next i
For i = 1 To 28
cop(i - 1).Picture = hole.Picture
Next i
m1.Picture = ehole.Picture
m2.Picture = ehole.Picture
m3.Picture = ehole.Picture
m4.Picture = ehole.Picture
you1.Picture = ehole.Picture
you2.Picture = ehole.Picture
you3.Picture = ehole.Picture
you4.Picture = ehole.Picture
For i = 1 To 4
FNum = i
FileName = "boardm" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Input #FNum, TextData
sp = TextData
Close
spot(sp).Picture = bm.Picture
End If
Next i
For i = 1 To 4
FNum = i
FileName = "boardc" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Input #FNum, TextData
sp = TextData
Close
spot(sp).Picture = rm.Picture
End If
Next i
For i = 1 To 4
FNum = i
FileName = "cop" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Input #FNum, TextData
sp = TextData
Close
cop(sp).Picture = rm.Picture
End If
Next i
For i = 1 To 4
FNum = i
FileName = "homem" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Input #FNum, TextData
sp = TextData
Close
End If
If FNum = 1 And sp = 1 Then
m1.Picture = mhome.Picture
ElseIf FNum = 2 And sp = 2 Then
m2.Picture = mhome.Picture
ElseIf FNum = 3 And sp = 3 Then
m3.Picture = mhome.Picture
ElseIf FNum = 4 And sp = 4 Then
m4.Picture = mhome.Picture
Else
End If
Next i
For i = 1 To 4
FNum = i
FileName = "homec" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Input #FNum, TextData
sp = TextData
Close
End If
If FNum = 1 And sp = 5 Then
you1.Picture = chome.Picture
ElseIf FNum = 2 And sp = 6 Then
you2.Picture = chome.Picture
ElseIf FNum = 3 And sp = 7 Then
you3.Picture = chome.Picture
ElseIf FNum = 4 And sp = 8 Then
you4.Picture = chome.Picture
Else
End If
Next i
csg = True
Exit Sub
ErrorHandler:
Resume Next
End Sub

Private Sub mnupass_Click()
If Timer6.Enabled = True Then Exit Sub
If player = False Then Exit Sub
ilpass = True
Timer2.Enabled = False
player = False
spin = False
finish = False
computer = True
mblink.Visible = False
yblink.Visible = True
Timer4.Enabled = True
Command1.Enabled = False
cspin.Enabled = True
cspin.SetFocus
mn.Tag = 0
mn.Caption = 0
stuck = False
icant.Visible = False
Timer9.Enabled = False
Timer10.Enabled = False
mp.Visible = False
my.Visible = False
mnueasy.Enabled = False
mnudiff.Enabled = False
mnuhard.Enabled = False
mnublue.Enabled = False
mnubluec.Enabled = False
mnugreen.Enabled = False
mnugreenc.Enabled = False
mnupink.Enabled = False
mnupinkc.Enabled = False
mnured.Enabled = False
mnuredc.Enabled = False
mnuwhite.Enabled = False
mnuwhitec.Enabled = False
mnuyellow.Enabled = False
mnuyellowc.Enabled = False
Call cspin_Click
End Sub

Private Sub mnupink_Click()
If mnupinkc.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
mn.ForeColor = QBColor(13)
mnublue.Checked = False
mnugreen.Checked = False
mnupink.Checked = True
mnured.Checked = False
mnuwhite.Checked = False
mnuyellow.Checked = False
dragm.Picture = pball.Picture
For i = 1 To 32
If spot(i - 1).Picture = bm.Picture Then
spot(i - 1).Picture = pkm.Picture
Else
End If
Next i
bm.Picture = pkm.Picture
If m1.Picture <> ehole.Picture Then
m1.Picture = bpm.Picture
Else
End If
If m2.Picture <> ehole.Picture Then
m2.Picture = bpm.Picture
Else
End If
If m3.Picture <> ehole.Picture Then
m3.Picture = bpm.Picture
Else
End If
If m4.Picture <> ehole.Picture Then
m4.Picture = bpm.Picture
Else
End If
mhome.Picture = bpm.Picture
blue = False
green = False
pink = True
red = False
white = False
yellow = False
End Sub


Private Sub mnupinkc_Click()
If mnupink.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
yn.ForeColor = QBColor(13)
mnubluec.Checked = False
mnugreenc.Checked = False
mnupinkc.Checked = True
mnuredc.Checked = False
mnuwhitec.Checked = False
mnuyellowc.Checked = False
For i = 1 To 32
If spot(i - 1).Picture = rm.Picture Then
spot(i - 1).Picture = pkm.Picture
Else
End If
Next i
rm.Picture = pkm.Picture
If you1.Picture <> ehole.Picture Then
you1.Picture = bpm.Picture
Else
End If
If you2.Picture <> ehole.Picture Then
you2.Picture = bpm.Picture
Else
End If
If you3.Picture <> ehole.Picture Then
you3.Picture = bpm.Picture
Else
End If
If you4.Picture <> ehole.Picture Then
you4.Picture = bpm.Picture
Else
End If
chome.Picture = bpm.Picture
comblue = False
comgreen = False
compink = True
comred = False
comwhite = False
comyellow = False
End Sub


Private Sub mnuquit_Click()
Load Form2
Form2.Show 1
End Sub


Private Sub mnured_Click()
If mnuredc.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
mn.ForeColor = QBColor(12)
mnublue.Checked = False
mnugreen.Checked = False
mnupink.Checked = False
mnured.Checked = True
mnuwhite.Checked = False
mnuyellow.Checked = False
dragm.Picture = rball.Picture
For i = 1 To 32
If spot(i - 1).Picture = bm.Picture Then
spot(i - 1).Picture = rdm.Picture
Else
End If
Next i
bm.Picture = rdm.Picture
If m1.Picture <> ehole.Picture Then
m1.Picture = brm.Picture
Else
End If
If m2.Picture <> ehole.Picture Then
m2.Picture = brm.Picture
Else
End If
If m3.Picture <> ehole.Picture Then
m3.Picture = brm.Picture
Else
End If
If m4.Picture <> ehole.Picture Then
m4.Picture = brm.Picture
Else
End If
mhome.Picture = brm.Picture
blue = False
green = False
pink = False
red = True
white = False
yellow = False
End Sub

Private Sub mnuredc_Click()
If mnured.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
yn.ForeColor = QBColor(12)
mnubluec.Checked = False
mnugreenc.Checked = False
mnupinkc.Checked = False
mnuredc.Checked = True
mnuwhitec.Checked = False
mnuyellowc.Checked = False
For i = 1 To 32
If spot(i - 1).Picture = rm.Picture Then
spot(i - 1).Picture = rdm.Picture
Else
End If
Next i
rm.Picture = rdm.Picture
If you1.Picture <> ehole.Picture Then
you1.Picture = brm.Picture
Else
End If
If you2.Picture <> ehole.Picture Then
you2.Picture = brm.Picture
Else
End If
If you3.Picture <> ehole.Picture Then
you3.Picture = brm.Picture
Else
End If
If you4.Picture <> ehole.Picture Then
you4.Picture = brm.Picture
Else
End If
chome.Picture = brm.Picture
comblue = False
comgreen = False
compink = False
comred = True
comwhite = False
comyellow = False
End Sub


Private Sub mnuresc_Click()
ress = True
mscore.Caption = 0
yscore.Caption = 0
End Sub

Private Sub mnusaveg_Click()
If finish = True Then Exit Sub
Dim FName, FNum, TestString
FNum = 1
TestString = 1
FName = "saveg" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
FNum = 1
FName = "levlg" & FNum & ".dat"
Open FName For Output As FNum
If mnueasy.Checked = True Then
TestString = 1
Write #FNum, TestString
ElseIf mnudiff.Checked = True Then
TestString = 2
Write #FNum, TestString
ElseIf mnuhard.Checked = True Then
TestString = 3
Write #FNum, TestString
Else
End If
Close
FNum = 1              'score
TestString = mscore.Caption
FName = "scorc" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
FNum = 1              'score
TestString = yscore.Caption
FName = "scorm" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
On Error GoTo ErrorHandler
Kill "homem1.dat"
Kill "homem2.dat"
Kill "homem3.dat"
Kill "homem4.dat"
Kill "homec1.dat"
Kill "homec2.dat"
Kill "homec3.dat"
Kill "homec4.dat"
Kill "boardm1.dat"
Kill "boardm2.dat"
Kill "boardm3.dat"
Kill "boardm4.dat"
Kill "boardc1.dat"
Kill "boardc2.dat"
Kill "boardc3.dat"
Kill "boardc4.dat"
Kill "cop1.dat"
Kill "cop2.dat"
Kill "cop3.dat"
Kill "cop4.dat"
FNum = 0
For i = 1 To 28
If spot(i - 1).Picture = bm.Picture Then
TestString = i - 1
FNum = FNum + 1
FName = "boardm" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
Next i
FNum = 0
For i = 1 To 32
If spot(i - 1).Picture = rm.Picture Then
TestString = i - 1
FNum = FNum + 1
FName = "boardc" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
Next i
FNum = 0
For i = 1 To 28
If cop(i - 1).Picture = rm.Picture Then
TestString = i - 1
FNum = FNum + 1
FName = "cop" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
Next i
If m1.Picture = mhome.Picture Then
TestString = 1
FNum = 1
FName = "homem" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
If m2.Picture = mhome.Picture Then
TestString = 2
FNum = 2
FName = "homem" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
If m3.Picture = mhome.Picture Then
TestString = 3
FNum = 3
FName = "homem" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
If m4.Picture = mhome.Picture Then
TestString = 4
FNum = 4
FName = "homem" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
If you1.Picture = chome.Picture Then
TestString = 5
FNum = 1
FName = "homec" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
If you2.Picture = chome.Picture Then
TestString = 6
FNum = 2
FName = "homec" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
If you3.Picture = chome.Picture Then
TestString = 7
FNum = 3
FName = "homec" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
If you4.Picture = chome.Picture Then
TestString = 8
FNum = 4
FName = "homec" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
End If
MsgBox "Your game has been saved successfully."
Exit Sub
ErrorHandler:
Resume Next
End Sub

Private Sub mnutake_Click()
If Timer6.Enabled = True Then Exit Sub
If ilpass = True Then Exit Sub
If cstuck = True Or stuck = True Then Exit Sub
If computer = True And cspin.Enabled = False Then Exit Sub
If hit = True Then
MsgBox "Sorry, if you take one of MY MARBLES!;  I'm not going to let you take it back.", 48, "He...he...he..."
Exit Sub
Else
End If
If player = True Then
MsgBox "You haven't moved yet!", 48, "Another chance..."
Exit Sub
Else
End If
mn.Caption = Image3.Tag
mn.Tag = Image3.Tag
If Image1.Tag = 100 Then
m1.Picture = mhome.Picture
GoTo 400
ElseIf Image1.Tag = 200 Then
m2.Picture = mhome.Picture
GoTo 400
ElseIf Image1.Tag = 300 Then
m3.Picture = mhome.Picture
GoTo 400
ElseIf Image1.Tag = 400 Then
m4.Picture = mhome.Picture
GoTo 400
Else
End If
spot(Image1.Tag).Picture = bm.Picture
400
spot(Image2.Tag).Picture = hole.Picture
Timer4.Enabled = False
player = True
computer = False
spin = True
yblink.Visible = False
Beep
Form1.MousePointer = 99
cspin.Enabled = False
Timer2.Enabled = True
mblink.Visible = True
End Sub

Private Sub mnuwhite_Click()
If mnuwhitec.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
mn.ForeColor = QBColor(15)
mnublue.Checked = False
mnugreen.Checked = False
mnupink.Checked = False
mnured.Checked = False
mnuwhite.Checked = True
mnuyellow.Checked = False
dragm.Picture = wball.Picture
For i = 1 To 32
If spot(i - 1).Picture = bm.Picture Then
spot(i - 1).Picture = wm.Picture
Else
End If
Next i
bm.Picture = wm.Picture
If m1.Picture <> ehole.Picture Then
m1.Picture = bwm.Picture
Else
End If
If m2.Picture <> ehole.Picture Then
m2.Picture = bwm.Picture
Else
End If
If m3.Picture <> ehole.Picture Then
m3.Picture = bwm.Picture
Else
End If
If m4.Picture <> ehole.Picture Then
m4.Picture = bwm.Picture
Else
End If
mhome.Picture = bwm.Picture
blue = False
green = False
pink = False
red = False
white = True
yellow = False
End Sub


Private Sub mnuwhitec_Click()
If mnuwhite.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
yn.ForeColor = QBColor(15)
mnubluec.Checked = False
mnugreenc.Checked = False
mnupinkc.Checked = False
mnuredc.Checked = False
mnuwhitec.Checked = True
mnuyellowc.Checked = False
For i = 1 To 32
If spot(i - 1).Picture = rm.Picture Then
spot(i - 1).Picture = wm.Picture
Else
End If
Next i
rm.Picture = wm.Picture
If you1.Picture <> ehole.Picture Then
you1.Picture = bwm.Picture
Else
End If
If you2.Picture <> ehole.Picture Then
you2.Picture = bwm.Picture
Else
End If
If you3.Picture <> ehole.Picture Then
you3.Picture = bwm.Picture
Else
End If
If you4.Picture <> ehole.Picture Then
you4.Picture = bwm.Picture
Else
End If
chome.Picture = bwm.Picture
comblue = False
comgreen = False
compink = False
comred = False
comwhite = True
comyellow = False
End Sub


Private Sub mnuyellow_Click()
If mnuyellowc.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
mn.ForeColor = QBColor(14)
mnublue.Checked = False
mnugreen.Checked = False
mnupink.Checked = False
mnured.Checked = False
mnuwhite.Checked = False
mnuyellow.Checked = True
dragm.Picture = yball.Picture
For i = 1 To 32
If spot(i - 1).Picture = bm.Picture Then
spot(i - 1).Picture = ym.Picture
Else
End If
Next i
bm.Picture = ym.Picture
If m1.Picture <> ehole.Picture Then
m1.Picture = bym.Picture
Else
End If
If m2.Picture <> ehole.Picture Then
m2.Picture = bym.Picture
Else
End If
If m3.Picture <> ehole.Picture Then
m3.Picture = bym.Picture
Else
End If
If m4.Picture <> ehole.Picture Then
m4.Picture = bym.Picture
Else
End If
mhome.Picture = bym.Picture
blue = False
green = False
pink = False
red = False
white = False
yellow = True
End Sub


Private Sub mnuyellowc_Click()
If mnuyellow.Checked = True Then
MsgBox "We both can't have the same color marbles!", 48
Exit Sub
Else
End If
yn.ForeColor = QBColor(14)
mnubluec.Checked = False
mnugreenc.Checked = False
mnupinkc.Checked = False
mnuredc.Checked = False
mnuwhitec.Checked = False
mnuyellowc.Checked = True
For i = 1 To 32
If spot(i - 1).Picture = rm.Picture Then
spot(i - 1).Picture = ym.Picture
Else
End If
Next i
rm.Picture = ym.Picture
If you1.Picture <> ehole.Picture Then
you1.Picture = bym.Picture
Else
End If
If you2.Picture <> ehole.Picture Then
you2.Picture = bym.Picture
Else
End If
If you3.Picture <> ehole.Picture Then
you3.Picture = bym.Picture
Else
End If
If you4.Picture <> ehole.Picture Then
you4.Picture = bym.Picture
Else
End If
chome.Picture = bym.Picture
comblue = False
comgreen = False
compink = False
comred = False
comwhite = False
comyellow = True
End Sub


Private Sub mscore_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 109
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
109
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub









Private Sub spot_DblClick(Index As Integer)
If player = False Or spin = False Or spot(Index).Picture = rm.Picture Or spot(Index).Picture = hole.Picture Then Exit Sub
MsgBox "Don't click me...hold down the left mouse button to pick me up, then drag me over to where you want to move, then release the mouse button."
End Sub

Private Sub spot_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
n = mn.Caption
If spot(Index).Picture = rm.Picture And pickup = False Then
GoTo 32
ElseIf spot(Index).Picture = rm.Picture And pickup = True Then
GoTo 33
Else
GoTo 34
End If
32
k = spot(Index).Tag
If home.Picture = mhome.Picture Then
If spot(Index).Index = n - 1 Then
spot(Index).Picture = bm.Picture
cop(k - 1).Picture = hole.Picture
If home.Tag = 1 Then
m1.Picture = ehole.Picture
ElseIf home.Tag = 2 Then
m2.Picture = ehole.Picture
ElseIf home.Tag = 3 Then
m3.Picture = ehole.Picture
ElseIf home.Tag = 4 Then
m4.Picture = ehole.Picture
Else
End If
Call gotyou
Timer2.Enabled = False
player = False
spin = False
computer = True
mblink.Visible = False
yblink.Visible = True
Timer4.Enabled = True
cspin.Enabled = True
cspin.SetFocus
mn.Tag = 0
mn.Caption = 0
stuck = False
hit = True
If six = True Then Call roll
Exit Sub
Else
If home.Tag = 1 Then
m1.Picture = home.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = home.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = home.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = home.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
End If
Exit Sub
33
k = spot(Index).Tag
If n0 = True Then
p = 0
ElseIf n1 = True Then
p = 1
ElseIf n2 = True Then
p = 2
ElseIf n3 = True Then
p = 3
ElseIf n4 = True Then
p = 4
ElseIf n5 = True Then
p = 5
ElseIf n6 = True Then
p = 6
ElseIf n7 = True Then
p = 7
ElseIf n8 = True Then
p = 8
ElseIf n9 = True Then
p = 9
ElseIf n10 = True Then
p = 10
ElseIf n11 = True Then
p = 11
ElseIf n12 = True Then
p = 12
ElseIf n13 = True Then
p = 13
ElseIf n14 = True Then
p = 14
ElseIf n15 = True Then
p = 15
ElseIf n16 = True Then
p = 16
ElseIf n17 = True Then
p = 17
ElseIf n18 = True Then
p = 18
ElseIf n19 = True Then
p = 19
ElseIf n20 = True Then
p = 20
ElseIf n21 = True Then
p = 21
ElseIf n22 = True Then
p = 22
ElseIf n23 = True Then
p = 23
ElseIf n24 = True Then
p = 24
ElseIf n25 = True Then
p = 25
ElseIf n26 = True Then
p = 26
Else
End If
If p + n = Index And p + n < 28 Then
spot(Index).Picture = bm.Picture
cop(k - 1).Picture = hole.Picture
pickup = False
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
Call gotyou
Timer2.Enabled = False
player = False
spin = False
computer = True
mblink.Visible = False
yblink.Visible = True
Timer4.Enabled = True
cspin.Enabled = True
cspin.SetFocus
mn.Tag = 0
mn.Caption = 0
stuck = False
hit = True
If six = True Then Call roll
Exit Sub
Else
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End If
Exit Sub
34
If spot(Index).Picture = bm.Picture And pickup = False Then GoTo 1
If spot(Index).Picture = bm.Picture And pickup = True Then GoTo 31
If pickup = True Then GoTo 30
If home.Picture = mhome.Picture Then
If spot(Index).Index = n - 1 Then
spot(Index).Picture = bm.Picture
If home.Tag = 1 Then
m1.Picture = ehole.Picture
ElseIf home.Tag = 2 Then
m2.Picture = ehole.Picture
ElseIf home.Tag = 3 Then
m3.Picture = ehole.Picture
ElseIf home.Tag = 4 Then
m4.Picture = ehole.Picture
Else
End If
Timer2.Enabled = False
player = False
spin = False
computer = True
mblink.Visible = False
yblink.Visible = True
Timer4.Enabled = True
cspin.Enabled = True
cspin.SetFocus
mn.Tag = 0
mn.Caption = 0
stuck = False
Image2.Tag = Index
If six = True Then Call roll
Exit Sub
Else
1
If home.Tag = 1 Then
m1.Picture = home.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = home.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = home.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = home.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
End If
Exit Sub
30
If n0 = True Then
p = 0
ElseIf n1 = True Then
p = 1
ElseIf n2 = True Then
p = 2
ElseIf n3 = True Then
p = 3
ElseIf n4 = True Then
p = 4
ElseIf n5 = True Then
p = 5
ElseIf n6 = True Then
p = 6
ElseIf n7 = True Then
p = 7
ElseIf n8 = True Then
p = 8
ElseIf n9 = True Then
p = 9
ElseIf n10 = True Then
p = 10
ElseIf n11 = True Then
p = 11
ElseIf n12 = True Then
p = 12
ElseIf n13 = True Then
p = 13
ElseIf n14 = True Then
p = 14
ElseIf n15 = True Then
p = 15
ElseIf n16 = True Then
p = 16
ElseIf n17 = True Then
p = 17
ElseIf n18 = True Then
p = 18
ElseIf n19 = True Then
p = 19
ElseIf n20 = True Then
p = 20
ElseIf n21 = True Then
p = 21
ElseIf n22 = True Then
p = 22
ElseIf n23 = True Then
p = 23
ElseIf n24 = True Then
p = 24
ElseIf n25 = True Then
p = 25
ElseIf n26 = True Then
p = 26
Else
End If
If p + n = Index And p + n < 28 Then
spot(Index).Picture = bm.Picture
pickup = False
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
Timer2.Enabled = False
player = False
spin = False
computer = True
mblink.Visible = False
yblink.Visible = True
Timer4.Enabled = True
cspin.Enabled = True
cspin.SetFocus
mn.Tag = 0
mn.Caption = 0
stuck = False
Image2.Tag = Index
Call win
If six = True Then Call roll
Else
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End If
Exit Sub
31
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub spot_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If hint = True Then Exit Sub
If spin = False Then Exit Sub
If player = False Then Exit Sub
If spot(Index).Picture = hole.Picture Then Exit Sub
If spot(Index).Picture = rm.Picture Then Exit Sub
spot(Index).DragIcon = dragm.Picture
spot(Index).Drag
pickup = True
spot(Index).Picture = hole.Picture
If Index = 0 Then
n0 = True
ElseIf Index = 1 Then
n1 = True
ElseIf Index = 2 Then
n2 = True
ElseIf Index = 3 Then
n3 = True
ElseIf Index = 4 Then
n4 = True
ElseIf Index = 5 Then
n5 = True
ElseIf Index = 6 Then
n6 = True
ElseIf Index = 7 Then
n7 = True
ElseIf Index = 8 Then
n8 = True
ElseIf Index = 9 Then
n9 = True
ElseIf Index = 10 Then
n10 = True
ElseIf Index = 11 Then
n11 = True
ElseIf Index = 12 Then
n12 = True
ElseIf Index = 13 Then
n13 = True
ElseIf Index = 14 Then
n14 = True
ElseIf Index = 15 Then
n15 = True
ElseIf Index = 16 Then
n16 = True
ElseIf Index = 17 Then
n17 = True
ElseIf Index = 18 Then
n18 = True
ElseIf Index = 19 Then
n19 = True
ElseIf Index = 20 Then
n20 = True
ElseIf Index = 21 Then
n21 = True
ElseIf Index = 22 Then
n22 = True
ElseIf Index = 23 Then
n23 = True
ElseIf Index = 24 Then
n24 = True
ElseIf Index = 25 Then
n25 = True
ElseIf Index = 26 Then
n26 = True
Else
End If
start = True
Image1.Tag = Index
Image3.Tag = mn.Tag
ilpass = False
End Sub


Private Sub Timer1_Timer()
Label1.Caption = Format(Time, "hh:mm AM/PM")
End Sub


Private Sub Timer10_Timer()
Select Case motsix
Case 1
mp.Move mp.Left - 5, mp.Top - 5
If mp.Left <= 522 Then
motsix = 2
ElseIf mp.Top <= 10 Then
motsix = 4
End If
Case 2
mp.Move mp.Left + 5, mp.Top - 5
If mp.Left >= 608 Then
motsix = 1
ElseIf mp.Top <= 10 Then
motsix = 3
End If
Case 3
mp.Move mp.Left + 5, mp.Top + 5
If mp.Left >= 608 Then
motsix = 4
ElseIf mp.Top >= 405 Then
motsix = 2
End If
Case 4
mp.Move mp.Left - 5, mp.Top + 5
If mp.Left <= 522 Then
motsix = 3
ElseIf mp.Top >= 405 Then
motsix = 1
End If
End Select
End Sub









Private Sub Timer11_Timer()
Static cml As Integer
Static cmr As Integer
cml = cml + 1
cmr = cmr + 1
If cml = 1 Then
my.Picture = mb.Picture
ElseIf cml = 2 Then
my.Picture = mg.Picture
ElseIf cml = 3 Then
my.Picture = mpp.Picture
ElseIf cml = 4 Then
my.Picture = mr.Picture
ElseIf cml = 5 Then
my.Picture = mw.Picture
ElseIf cml = 6 Then
my.Picture = myy.Picture
Else
End If
If cmr = 6 Then
mp.Picture = mb.Picture
ElseIf cmr = 5 Then
mp.Picture = mg.Picture
ElseIf cmr = 4 Then
mp.Picture = mpp.Picture
ElseIf cmr = 3 Then
mp.Picture = mr.Picture
ElseIf cmr = 2 Then
mp.Picture = mw.Picture
ElseIf cmr = 1 Then
mp.Picture = myy.Picture
Else
End If
If cml = 6 Then
cml = 0
cmr = 0
Else
End If
End Sub

Private Sub Timer12_Timer()
Static h As Integer
Static k As Integer
Static o As Integer
If newgame = True Then
h = 0
k = 0
o = 0
loc1 = False
loc2 = False
loc3 = False
loc4 = False
have = False
newgame = False
foundhit = False
foundhome = False
foundhole = False
noplace = False
hint = False
Exit Sub
Else
End If
If m1.Picture <> mhome.Picture And m2.Picture <> mhome.Picture And m3.Picture <> mhome.Picture And m4.Picture <> mhome.Picture Then
have = False
Else
have = True
End If
If foundhit = True Then
GoTo 190
Else
End If
If foundhome = True Then
GoTo 201
Else
End If
If foundhole = True Then
GoTo 195
Else
End If
If noplace = True Then
GoTo 301
Else
End If
For i = 1 To 23
If spot(i - 1).Picture = bm.Picture And spot(i - 1 + mn.Tag).Picture = rm.Picture And spot(i - 1 + mn.Tag).Index < 24 Then
foundhit = True
o = i
GoTo 190
Else
End If
Next i
For i = 1 To 6
If spot(i - 1).Picture = rm.Picture And i = mn.Tag And have = True Then
foundhome = True
o = i
GoTo 200
Else
End If
Next i
For i = 1 To 27
If spot(i - 1).Picture = bm.Picture And spot(i - 1 + mn.Tag).Picture = hole.Picture And spot(i - 1 + mn.Tag).Index < 28 Then
foundhole = True
o = i
GoTo 195
Else
End If
Next i
For i = 1 To 6
If spot(i - 1).Picture = hole.Picture And i = mn.Tag And have = True Then
noplace = True
o = i
GoTo 300
Else
End If
Next i
190
If h > 2 Then GoTo 192
If k = 1 Then GoTo 191
h = h + 1
If blue = True Then
spot(o - 1).Picture = whiteb.Picture
ElseIf green = True Then
spot(o - 1).Picture = whiteg.Picture
ElseIf pink = True Then
spot(o - 1).Picture = whitep.Picture
ElseIf red = True Then
spot(o - 1).Picture = whiter.Picture
ElseIf white = True Then
spot(o - 1).Picture = whitew.Picture
ElseIf yellow = True Then
spot(o - 1).Picture = whitey.Picture
Else
End If
k = 1
Exit Sub
191
spot(o - 1).Picture = bm.Picture
k = 0
Exit Sub
192
spot(o - 1).Picture = bm.Picture
If k = 1 Then GoTo 193
h = h + 1
If comblue = True Then
spot(o - 1 + mn.Tag).Picture = whiteb.Picture
ElseIf comgreen = True Then
spot(o - 1 + mn.Tag).Picture = whiteg.Picture
ElseIf compink = True Then
spot(o - 1 + mn.Tag).Picture = whitep.Picture
ElseIf comred = True Then
spot(o - 1 + mn.Tag).Picture = whiter.Picture
ElseIf comwhite = True Then
spot(o - 1 + mn.Tag).Picture = whitew.Picture
ElseIf comyellow = True Then
spot(o - 1 + mn.Tag).Picture = whitey.Picture
Else
End If
k = 1
Exit Sub
193
spot(o - 1 + mn.Tag).Picture = rm.Picture
k = 0
If h = 6 Then
h = 0
foundhit = False
hint = False
Timer12.Enabled = False
Exit Sub
Else
End If
Exit Sub
200
If m1.Picture <> ehole.Picture Then
loc1 = True
foundhome = True
GoTo 201
ElseIf m2.Picture <> ehole.Picture Then
loc2 = True
foundhome = True
GoTo 201
ElseIf m3.Picture <> ehole.Picture Then
loc3 = True
foundhome = True
GoTo 201
ElseIf m4.Picture <> ehole.Picture Then
loc4 = True
foundhome = True
GoTo 201
Else
End If
201
If h > 5 Then GoTo 203
If k = 1 Then GoTo 202
h = h + 1
If loc1 = True Then
m1.Picture = ehole.Picture
k = 1
Exit Sub
ElseIf loc2 = True Then
m2.Picture = ehole.Picture
k = 1
Exit Sub
ElseIf loc3 = True Then
m3.Picture = ehole.Picture
k = 1
Exit Sub
ElseIf loc4 = True Then
m4.Picture = ehole.Picture
k = 1
Exit Sub
Else
End If
202
h = h + 1
If loc1 = True Then
m1.Picture = mhome.Picture
k = 0
ElseIf loc2 = True Then
m2.Picture = mhome.Picture
k = 0
ElseIf loc3 = True Then
m3.Picture = mhome.Picture
k = 0
ElseIf loc4 = True Then
m4.Picture = mhome.Picture
k = 0
Else
End If
Exit Sub
203
If k = 1 Then GoTo 194
h = h + 1
If comblue = True Then
spot(o - 1).Picture = whiteb.Picture
ElseIf comgreen = True Then
spot(o - 1).Picture = whiteg.Picture
ElseIf compink = True Then
spot(o - 1).Picture = whitep.Picture
ElseIf comred = True Then
spot(o - 1).Picture = whiter.Picture
ElseIf comwhite = True Then
spot(o - 1).Picture = whitew.Picture
ElseIf comyellow = True Then
spot(o - 1).Picture = whitey.Picture
Else
End If
k = 1
Exit Sub
194
spot(o - 1).Picture = rm.Picture
k = 0
If h = 9 Then
h = 0
loc1 = False
loc2 = False
loc3 = False
loc4 = False
foundhome = False
hint = False
have = False
Timer12.Enabled = False
Exit Sub
Else
End If
Exit Sub
195
If h > 2 Then GoTo 197
If k = 1 Then GoTo 196
h = h + 1
If blue = True Then
spot(o - 1).Picture = whiteb.Picture
ElseIf green = True Then
spot(o - 1).Picture = whiteg.Picture
ElseIf pink = True Then
spot(o - 1).Picture = whitep.Picture
ElseIf red = True Then
spot(o - 1).Picture = whiter.Picture
ElseIf white = True Then
spot(o - 1).Picture = whitew.Picture
ElseIf yellow = True Then
spot(o - 1).Picture = whitey.Picture
Else
End If
k = 1
Exit Sub
196
spot(o - 1).Picture = bm.Picture
k = 0
Exit Sub
197
spot(o - 1).Picture = bm.Picture
If k = 1 Then GoTo 198
h = h + 1
spot(o - 1 + mn.Tag).Picture = whiteh.Picture
k = 1
Exit Sub
198
spot(o - 1 + mn.Tag).Picture = hole.Picture
k = 0
If h = 6 Then
h = 0
foundhole = False
hint = False
Timer12.Enabled = False
Exit Sub
Else
End If
Exit Sub
300
If m1.Picture <> ehole.Picture Then
loc1 = True
noplace = True
GoTo 301
ElseIf m2.Picture <> ehole.Picture Then
loc2 = True
noplace = True
GoTo 301
ElseIf m3.Picture <> ehole.Picture Then
loc3 = True
noplace = True
GoTo 301
ElseIf m4.Picture <> ehole.Picture Then
loc4 = True
noplace = True
GoTo 301
Else
End If
301
If h > 5 Then GoTo 303
If k = 1 Then GoTo 302
h = h + 1
If loc1 = True Then
m1.Picture = ehole.Picture
k = 1
Exit Sub
ElseIf loc2 = True Then
m2.Picture = ehole.Picture
k = 1
Exit Sub
ElseIf loc3 = True Then
m3.Picture = ehole.Picture
k = 1
Exit Sub
ElseIf loc4 = True Then
m4.Picture = ehole.Picture
k = 1
Exit Sub
Else
End If
302
h = h + 1
If loc1 = True Then
m1.Picture = mhome.Picture
k = 0
ElseIf loc2 = True Then
m2.Picture = mhome.Picture
k = 0
ElseIf loc3 = True Then
m3.Picture = mhome.Picture
k = 0
ElseIf loc4 = True Then
m4.Picture = mhome.Picture
k = 0
Else
End If
Exit Sub
303
If k = 1 Then GoTo 304
h = h + 1
spot(o - 1).Picture = whiteh.Picture
k = 1
Exit Sub
304
spot(o - 1).Picture = hole.Picture
k = 0
If h = 9 Then
h = 0
loc1 = False
loc2 = False
loc3 = False
loc4 = False
noplace = False
hint = False
have = False
Timer12.Enabled = False
Exit Sub
Else
End If
End Sub

Private Sub Timer2_Timer()
Static a As Integer
If a = 1 Then
GoTo 170
End If
mblink.Picture = Picture1.Picture
a = 1
Exit Sub
170
mblink.Picture = Picture2.Picture
a = 0
End Sub


Private Sub Timer3_Timer()
Static b
b = b + 1
Randomize
s = Int(Rnd * 6) + 1
mn.Caption = s
mn.Tag = s
If b = 20 Then
b = 0
Timer3.Enabled = False
Beep
spin = True
If mn.Tag = 6 Then
six = True
Else
six = False
End If
End If
End Sub


Private Sub Timer4_Timer()
Static b As Integer
If b = 1 Then
GoTo 171
End If
yblink.Picture = Picture1.Picture
b = 1
Exit Sub
171
yblink.Picture = Picture2.Picture
b = 0
End Sub


Private Sub evaluate()
If m1.Picture <> mhome.Picture Then
m1.Picture = mhome.Picture
GoTo 51
ElseIf m2.Picture <> mhome.Picture Then
m2.Picture = mhome.Picture
GoTo 51
ElseIf m3.Picture <> mhome.Picture Then
m3.Picture = mhome.Picture
GoTo 51
ElseIf m4.Picture <> mhome.Picture Then
m4.Picture = mhome.Picture
51
Else
End If
End Sub

Private Sub Timer6_Timer()
If computer = False Then GoTo 71
If youwin.Visible = True Then
youwin.Visible = False
ElseIf youwin.Visible = False Then
youwin.Visible = True
Else
End If
Exit Sub
71
If iwin.Visible = True Then
iwin.Visible = False
ElseIf iwin.Visible = False Then
iwin.Visible = True
Else
End If
End Sub


Private Sub Timer5_Timer()
Static c
c = c + 1
Randomize
s = Int(Rnd * 6) + 1
yn.Caption = s
yn.Tag = s
If c = 20 Then
c = 0
Timer5.Enabled = False
If yn.Tag = 6 Then
six = True
Else
six = False
End If
End If
End Sub



Private Sub gotyou()
If you1.Picture <> chome.Picture Then
you1.Picture = chome.Picture
GoTo 50
ElseIf you2.Picture <> chome.Picture Then
you2.Picture = chome.Picture
GoTo 50
ElseIf you3.Picture <> chome.Picture Then
you3.Picture = chome.Picture
GoTo 50
ElseIf you4.Picture <> chome.Picture Then
you4.Picture = chome.Picture
50
Else
End If
End Sub

Private Sub Timer7_Timer()
If mn.Tag = 0 Then Exit Sub
If Timer3.Enabled = True Or start = True Then Exit Sub
u = 0
p = 0
s = 0
n = mn.Tag
For l = 1 To 27
If spot(l - 1).Picture = bm.Picture And spot(l - 1 + n).Picture = bm.Picture Then
u = u + 1
Else
End If
Next l
For m = 1 To 27
If spot(m).Picture = bm.Picture And spot(m + n).Index > 27 Then
p = p + 1
Else
End If
Next m
s = u + p
If s = 4 Then
stuck = True
six = False
youcant.Visible = True
Timer2.Enabled = False
mn.Tag = 0
player = False
spin = False
computer = True
mblink.Visible = False
yblink.Visible = True
Timer4.Enabled = True
cspin.Enabled = True
cspin.SetFocus
Else
End If
End Sub

Private Sub Timer8_Timer()
Timer8.Enabled = False
p = 0
f = 0
n = 0
e = yn.Tag
For i = 1 To 27
If cop(i - 1).Picture = rm.Picture And cop(i - 1 + e).Picture = rm.Picture Then
p = p + 1
Else
End If
Next i
For i = 1 To 27
If cop(i).Picture = rm.Picture And cop(i + e).Index > 27 Then
f = f + 1
Else
End If
Next i
n = p + f
If n = 4 Then
cstuck = True
six = False
spag.Visible = False
icant.Visible = True
Timer4.Enabled = False
player = True
computer = False
mblink.Visible = True
yblink.Visible = False
Beep
Form1.MousePointer = 99
Timer2.Enabled = True
Command1.Enabled = True
Command1.SetFocus
Timer8.Enabled = False
Exit Sub
Else
cstuck = False
Timer8.Enabled = False
End If
End Sub

Private Sub Timer9_Timer()
Select Case motone
Case 1
my.Move my.Left - 5, my.Top - 5
If my.Left <= 0 Then
motone = 2
ElseIf my.Top <= 10 Then
motone = 4
End If
Case 2
my.Move my.Left + 5, my.Top - 5
If my.Left >= 70 Then
motone = 1
ElseIf my.Top <= 10 Then
motone = 3
End If
Case 3
my.Move my.Left + 5, my.Top + 5
If my.Left >= 70 Then
motone = 4
ElseIf my.Top >= 405 Then
motone = 2
End If
Case 4
my.Move my.Left - 5, my.Top + 5
If my.Left <= 0 Then
motone = 3
ElseIf my.Top >= 405 Then
motone = 1
End If
End Select
End Sub

Private Sub yblink_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 113
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
113
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub yn_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 69
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
69
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub you1_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 60
If pickup = False Then
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Source.Picture = mhome.Picture
Exit Sub
Else
End If
60
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub you2_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 61
If pickup = False Then
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Source.Picture = mhome.Picture
Exit Sub
Else
End If
61
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub you3_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 62
If pickup = False Then
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Source.Picture = mhome.Picture
Exit Sub
Else
End If
62
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub


Private Sub you4_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 63
If pickup = False Then
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Source.Picture = mhome.Picture
Exit Sub
Else
End If
63
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub



Private Sub win()
Static mpoint As Integer
Static ypoint As Integer
If csg = False Then
GoTo 665
Else
End If
Dim FileName, TextData, FNum, FName, TestString
FNum = 1
FileName = "scorm" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Do While Not EOF(FNum)
Input #FNum, TextData
Loop
sp = TextData
ypoint = sp
Close
End If
FNum = 1
FileName = "scorc" & FNum & ".dat"
If Len(FileName) Then
Open FileName For Input As FNum
Do While Not EOF(FNum)
Input #FNum, TextData
Loop
sp = TextData
mpoint = sp
Close
End If
665
If spot(24).Picture = bm.Picture And spot(25).Picture = bm.Picture And spot(26).Picture = bm.Picture And spot(27).Picture = bm.Picture Then
Timer6.Enabled = True
Timer4.Enabled = False
player = False
sgl = True
spag.Visible = False
computer = True
mblink.Visible = False
yblink.Visible = False
Form1.MousePointer = 99
Timer2.Enabled = False
Command1.Enabled = False
cspin.Enabled = False
six = False
SoundName$ = "noise2.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
If ress = True Then
ypoint = 0
ress = False
Else
End If
ypoint = ypoint + 1
yscore.Caption = ypoint
finish = True
Timer9.Enabled = True
Timer10.Enabled = True
mp.Visible = True
my.Visible = True
If csg = True Then
FNum = 1              'score
TestString = yscore.Caption
FName = "scorm" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
Else
End If
Exit Sub
Else
End If
If spot(28).Picture = rm.Picture And spot(29).Picture = rm.Picture And spot(30).Picture = rm.Picture And spot(31).Picture = rm.Picture Then
Timer6.Enabled = True
Timer4.Enabled = False
player = False
sgl = True
spag.Visible = False
computer = False
mblink.Visible = False
yblink.Visible = False
Form1.MousePointer = 99
Timer2.Enabled = False
Command1.Enabled = False
cspin.Enabled = False
six = False
comwin = True
SoundName$ = "noise2.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
If ress = True Then
mpoint = 0
ress = False
Else
End If
mpoint = mpoint + 1
mscore.Caption = mpoint
finish = True
Timer9.Enabled = True
Timer10.Enabled = True
mp.Visible = True
my.Visible = True
If csg = True Then
FNum = 1              'score
TestString = mscore.Caption
FName = "scorc" & FNum & ".dat"
Open FName For Output As FNum
Write #FNum, TestString
Close
Else
End If
Exit Sub
Else
End If
End Sub

Private Sub roll()
Msg = "You got a 6 ! , want to spin again?"
Style = vbYesNo + vbExclamation + vbDefaultButton1
Title = "Home First!"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
Timer4.Enabled = False
player = True
computer = False
mblink.Visible = True
yblink.Visible = False
Form1.MousePointer = 99
Timer2.Enabled = True
cspin.Enabled = False
Command1.Enabled = True
Command1.SetFocus
six = False
Else
six = False
End If
End Sub

Private Sub comdiff()
90
youcant.Visible = False
start = False
Dim numb As Integer
Dim tnumb As Integer
Dim s As Integer
cspin.Enabled = False
Timer5.Enabled = True
Do
DoEvents
Loop Until Timer5.Enabled = False
Timer8.Enabled = True
Form1.MousePointer = 11
Do
DoEvents
Loop Until Timer8.Enabled = False
If cstuck = True Then Exit Sub
s = yn.Caption
If you1.Picture <> chome.Picture And you2.Picture <> chome.Picture And you3.Picture <> chome.Picture And you4.Picture <> chome.Picture Then
nomarb = True
GoTo 5
Else
nomarb = False
GoTo 3
End If
3
For i = 1 To 28
If spot(i - 1).Tag = yn.Caption Then
If spot(i - 1).Picture <> rm.Picture And spot(i - 1).Picture <> bm.Picture Then
spot(i - 1).Picture = rm.Picture
cop(s - 1).Picture = rm.Picture
GoTo 4
Else
GoTo 5
End If
End If
Next i
5
For i = 1 To 28
If spot(i - 1).Tag = yn.Caption Then
If spot(i - 1).Picture = bm.Picture And nomarb = False Then
spot(i - 1).Picture = rm.Picture
cop(s - 1).Picture = rm.Picture
Call evaluate
GoTo 4
Else
End If
End If
Next i
numb = Int(Rnd * 29) + 1
10
For i = 1 To 300
t = Int(Rnd * 32)
If spot(t).Tag = numb And spot(t).Picture = rm.Picture And spot(t).Tag + s < 29 Then
tnumb = spot(t).Tag + s
GoTo 16
End If
Next i
numb = Int(Rnd * 29) + 1
GoTo 10
16
For j = 1 To 32
If spot(j - 1).Tag = tnumb And spot(j - 1).Picture = rm.Picture Then
numb = numb - 1
GoTo 10
ElseIf spot(j - 1).Tag = tnumb And spot(j - 1).Picture <> rm.Picture Then
spot(t).Picture = hole.Picture
cop(numb - 1).Picture = hole.Picture
Exit For
Else
End If
Next j
11
For i = 1 To 300
t = Int(Rnd * 32)
If spot(t).Tag = tnumb And spot(t).Picture <> bm.Picture And spot(t).Picture <> rm.Picture Then
spot(t).Picture = rm.Picture
cop(numb - 1 + s).Picture = rm.Picture
GoTo 17
Else
End If
If spot(t).Tag = tnumb And spot(t).Picture <> rm.Picture And spot(t).Picture <> hole.Picture Then
spot(t).Picture = rm.Picture
cop(numb - 1 + s).Picture = rm.Picture
Call evaluate
Exit For
Else
End If
Next i
17
Call win
If comwin = True Then Exit Sub
If six = True Then
Beep
spag.Visible = True
GoTo 90
Else
spag.Visible = False
End If
Timer4.Enabled = False
player = True
computer = False
mblink.Visible = True
yblink.Visible = False
Beep
Form1.MousePointer = 99
Timer2.Enabled = True
Command1.Enabled = True
Command1.SetFocus
Exit Sub
4
If you1.Picture <> ehole.Picture Then
you1.Picture = ehole.Picture
ElseIf you2.Picture <> ehole.Picture Then
you2.Picture = ehole.Picture
ElseIf you3.Picture <> ehole.Picture Then
you3.Picture = ehole.Picture
ElseIf you4.Picture <> ehole.Picture Then
you4.Picture = ehole.Picture
Else
End If
If six = True Then
Beep
spag.Visible = True
GoTo 90
Else
spag.Visible = False
End If
Timer4.Enabled = False
player = True
computer = False
mblink.Visible = True
yblink.Visible = False
Beep
Form1.MousePointer = 99
Timer2.Enabled = True
Command1.Enabled = True
Command1.SetFocus
Exit Sub
End Sub

Private Sub comhard()
90
youcant.Visible = False
start = False
Dim numb As Integer
Dim tnumb As Integer
Dim s As Integer
cspin.Enabled = False
Timer5.Enabled = True
Do
DoEvents
Loop Until Timer5.Enabled = False
Timer8.Enabled = True
Form1.MousePointer = 11
Do
DoEvents
Loop Until Timer8.Enabled = False
If cstuck = True Then Exit Sub
s = yn.Caption
If you1.Picture <> chome.Picture And you2.Picture <> chome.Picture And you3.Picture <> chome.Picture And you4.Picture <> chome.Picture Then
nomarb = True
GoTo 5
Else
nomarb = False
GoTo 3
End If
3
For i = 1 To 28
If spot(i - 1).Tag = yn.Caption Then
If spot(i - 1).Picture <> rm.Picture And spot(i - 1).Picture <> bm.Picture Then
spot(i - 1).Picture = rm.Picture
cop(s - 1).Picture = rm.Picture
GoTo 4
Else
GoTo 5
End If
End If
Next i
5
For i = 1 To 28
If spot(i - 1).Tag = yn.Caption Then
If spot(i - 1).Picture = bm.Picture And nomarb = False Then
spot(i - 1).Picture = rm.Picture
cop(s - 1).Picture = rm.Picture
Call evaluate
GoTo 4
Else
End If
End If
Next i
numb = 28
10
If numb = 0 Then
numb = 28
done = True
Else
End If
For i = 1 To 300
t = Int(Rnd * 32)
If spot(t).Tag = numb And spot(t).Picture = rm.Picture And spot(t).Tag + s < 29 Then
tnumb = spot(t).Tag + s
GoTo 16
End If
Next i
numb = numb - 1
GoTo 10
16
For j = 1 To 32
If spot(j - 1).Tag = tnumb And spot(j - 1).Picture = rm.Picture Then
numb = numb - 1
GoTo 10
Else
End If
If spot(j - 1).Tag = tnumb And spot(j - 1).Picture <> rm.Picture And spot(j - 1).Picture <> hole.Picture Then
spot(t).Picture = hole.Picture
cop(numb - 1).Picture = hole.Picture
GoTo 11
Else
End If
If spot(j - 1).Tag = tnumb And spot(j - 1).Picture <> rm.Picture And spot(j - 1).Picture <> bm.Picture And done = True Then
spot(t).Picture = hole.Picture
cop(numb - 1).Picture = hole.Picture
GoTo 11
Else
End If
Next j
If numb = 0 Then
numb = 28
done = True
GoTo 10
Else
End If
numb = numb - 1
GoTo 10
11
done = False
For i = 1 To 300
t = Int(Rnd * 32)
If spot(t).Tag = tnumb And spot(t).Picture <> bm.Picture And spot(t).Picture <> rm.Picture Then
spot(t).Picture = rm.Picture
cop(numb - 1 + s).Picture = rm.Picture
GoTo 17
Else
End If
If spot(t).Tag = tnumb And spot(t).Picture <> rm.Picture And spot(t).Picture <> hole.Picture Then
spot(t).Picture = rm.Picture
cop(numb - 1 + s).Picture = rm.Picture
Call evaluate
Exit For
Else
End If
Next i
17
Call win
If comwin = True Then Exit Sub
If six = True Then
Beep
spag.Visible = True
GoTo 90
Else
spag.Visible = False
End If
Timer4.Enabled = False
player = True
computer = False
mblink.Visible = True
yblink.Visible = False
Beep
Form1.MousePointer = 99
Timer2.Enabled = True
Command1.Enabled = True
Command1.SetFocus
Exit Sub
4
If you1.Picture <> ehole.Picture Then
you1.Picture = ehole.Picture
ElseIf you2.Picture <> ehole.Picture Then
you2.Picture = ehole.Picture
ElseIf you3.Picture <> ehole.Picture Then
you3.Picture = ehole.Picture
ElseIf you4.Picture <> ehole.Picture Then
you4.Picture = ehole.Picture
Else
End If
If six = True Then
Beep
spag.Visible = True
GoTo 90
Else
spag.Visible = False
End If
Timer4.Enabled = False
player = True
computer = False
mblink.Visible = True
yblink.Visible = False
Beep
Form1.MousePointer = 99
Timer2.Enabled = True
Command1.Enabled = True
Command1.SetFocus
Exit Sub
End Sub

Private Sub yscore_DragDrop(Source As Control, X As Single, Y As Single)
If pickup = True Then GoTo 111
If pickup = False Then
If home.Tag = 1 Then
m1.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 2 Then
m2.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 3 Then
m3.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
ElseIf home.Tag = 4 Then
m4.Picture = mhome.Picture
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
Else
End If
End If
Exit Sub
111
Source.Picture = bm.Picture
pickup = False
SoundName$ = "noise1.wav"
wFlags% = SND_ASYNC Or SND_NODEFAULT
z% = sndPlaySound(SoundName$, wFlags%)
n0 = False
n1 = False
n2 = False
n3 = False
n4 = False
n5 = False
n6 = False
n7 = False
n8 = False
n9 = False
n10 = False
n11 = False
n12 = False
n13 = False
n14 = False
n15 = False
n16 = False
n17 = False
n18 = False
n19 = False
n20 = False
n21 = False
n22 = False
n23 = False
n24 = False
n25 = False
n26 = False
End Sub



Private Sub gamel()
If sgl = False Then Exit Sub
If mnueasy.Checked = True Then
leve.Visible = True
ElseIf mnudiff.Checked = True Then
levd.Visible = True
ElseIf mnuhard.Checked = True Then
levh.Visible = True
Else
End If
End Sub
