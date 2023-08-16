VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "form5"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form5.frx":0000
   Moveable        =   0   'False
   Picture         =   "Form5.frx":08CA
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Index           =   1
      Left            =   2040
      Picture         =   "Form5.frx":1A5DD
      ScaleHeight     =   735
      ScaleWidth      =   7440
      TabIndex        =   1
      Top             =   1680
      Width           =   7440
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   0
         X2              =   7440
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "QUIT"
      Height          =   615
      Left            =   5640
      MouseIcon       =   "Form5.frx":1BC29
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   360
      Top             =   360
   End
   Begin VB.Image Image1 
      Height          =   3255
      Index           =   0
      Left            =   3360
      Picture         =   "Form5.frx":1C4F3
      Top             =   2760
      Width           =   5280
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Private Sub ASD_Click()
Unload Me
Form7.Show
End Sub
Private Sub Command1_Click()
Unload Me
rs2.Show
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = True
Command1.Font.Size = 9
End Sub
Private Sub Command2_Click()
Unload Me
Form7.Show
End Sub
Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.Font.Bold = True
Command2.Font.Size = 9
End Sub
Private Sub Command3_Click()
Unload Me
Form8.Show
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.Font.Bold = True
Command3.Font.Size = 9

End Sub

Private Sub Command4_Click()
Unload Me
Form10.Show

End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.Font.Bold = True
Command4.Font.Size = 9

End Sub

Private Sub Command5_Click()
Unload Me
Form9.Show

End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.Font.Bold = True
Command5.Font.Size = 9

End Sub

Private Sub Command6_Click()
Unload Me
Form11.Show

End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.Font.Bold = True
Command6.Font.Size = 9

End Sub

Private Sub Command7_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command7.Font.Bold = True
Command7.Font.Size = 9

End Sub

Private Sub Timer1_Timer()
'Picture2.Picture = LoadPicture("c:\image0" & i & ".jpg")
i = i + 1
If i = 5 Then i = 0
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command7.Font.Bold = False
Command7.Font.Size = 8
End Sub

Private Sub VSD_Click()
Unload Me
Form8.Show
End Sub
