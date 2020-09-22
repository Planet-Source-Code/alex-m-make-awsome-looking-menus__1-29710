VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cool Menus"
   ClientHeight    =   3405
   ClientLeft      =   3210
   ClientTop       =   2970
   ClientWidth     =   5115
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label mnuexit 
      BackColor       =   &H00FF0000&
      Caption         =   " Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Lab3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":030A
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image menu 
      Height          =   2025
      Left            =   0
      Picture         =   "Form1.frx":0398
      Top             =   405
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Written by Alex M"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label lab1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Image but1 
      Height          =   525
      Left            =   0
      Picture         =   "Form1.frx":CC3E
      Top             =   -120
      Width           =   1335
   End
   Begin VB.Label Lab2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Image but2 
      Height          =   525
      Left            =   0
      Picture         =   "Form1.frx":F124
      Top             =   -120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!! Cool Menus - Written by Alex M in VB6 !!!!
'!!!! Here is a rather worthless program    !!!!
'!!!! unless you want cool menus in your    !!!!
'!!!! programs.  I hereby give permission to!!!!
'!!!! you (and I mean YOU) to use this code !!!!
'!!!! in your programs without giving me any!!!!
'!!!! credit what so ever.  I don't need it.!!!!
'!!!! Sorry I did not comment much, it is   !!!!
'!!!! all pretty simple and not complex.    !!!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Private Sub but1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'shows the mnu
but1.Visible = False
lab1.Visible = False
Lab3.Visible = True
menu.Visible = True
mnuexit.Visible = True
End Sub

Private Sub but1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.BackColor = &HFF0000 'exit button in menu = blue
End Sub

Private Sub but1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
but1.Visible = True
lab1.Visible = True
End Sub

Private Sub but2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.BackColor = &HFF0000 'exit button in menu = blue
End Sub

Private Sub Form_Click()
menu.Visible = False
Lab3.Visible = False
mnuexit.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.BackColor = &HFF0000 'exit button in menu = blue
End Sub

Private Sub lab1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
but1.Visible = False
lab1.Visible = False
Lab3.Visible = True
menu.Visible = True
mnuexit.Visible = True
End Sub

Private Sub lab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.BackColor = &HFF0000 'exit button in menu = blue
End Sub

Private Sub lab1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
but1.Visible = True
lab1.Visible = True
End Sub


Private Sub Lab2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.BackColor = &HFF0000 'exit button in menu = blue
End Sub

Private Sub Lab3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.BackColor = &HFF0000 'exit button in menu = blue
End Sub

Private Sub Label2_Click()
menu.Visible = False
Lab3.Visible = False
mnuexit.Visible = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.BackColor = &HFF0000 'exit button in menu = blue
End Sub

Private Sub menu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.BackColor = &HFF0000 'exit button in menu = blue
End Sub

Private Sub mnuexit_Click()
End ' exit program
End Sub

Private Sub mnuexit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.BackColor = &HFF00& 'exit button in menu = green
End Sub
