VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4875
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4875
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   840
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   0
      Picture         =   "home2.frx":0000
      Top             =   6480
      Width           =   9600
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Form1
Unload Me
End
End Sub

Private Sub Image1_Click()
Unload Form1
Unload Me
End
End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False
Do
Image1.Move Image1.Left + 0, Image1.Top - 10
If Image1.Top < 0 Then
Exit Do
End If
DoEvents
Loop
Call Sleep(1000)
Unload Form1
Unload Me
End
End Sub


