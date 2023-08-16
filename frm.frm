VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "WELCOME"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Picture         =   "frm.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   7800
      TabIndex        =   4
      Top             =   1200
      Width           =   7800
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "&CANCEL"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&OK"
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Timer tmrScroll 
      Left            =   2040
      Top             =   7560
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   7560
   End
   Begin VB.PictureBox picHolder 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   3240
      ScaleHeight     =   1440
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   5880
      Width           =   5055
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   $"frm.frx":164C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1335
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Image Image1 
      Height          =   3345
      Left            =   3600
      Picture         =   "frm.frx":171A
      Top             =   2280
      Width           =   4500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form1
Form2.Show

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = True
Command1.Font.Size = 9
End Sub

Private Sub Command2_Click()
Unload Form1

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.Font.Bold = True
Command2.Font.Size = 9

End Sub

Private Sub Form_Load()
'Pause Feature by _ChAdWiCk_
   lblMsg.Top = picHolder.Height
   tmrScroll.Interval = 1
   tmrScroll.Enabled = True
   
'WebBrowser1.Navigate "c:\mainabc.gif"
'WebBrowser1.Navigate "about:<html><body bgcolor=brown scroll='no'><img src='c:\mainabc.gif'></img></body></html>"
End Sub
Private Sub Timer1_Timer()
tmrScroll.Enabled = True
Timer1.Enabled = False

End Sub
Private Sub tmrScroll_Timer()
If lblMsg.Top > -lblMsg.Height Then
       lblMsg.Top = lblMsg.Top - 10
           If lblMsg.Top = 255 Then
           pause
           End If
   Else
       lblMsg.Top = picHolder.Height
   End If
End Sub

Sub pause()
tmrScroll.Enabled = False
Timer1.Interval = 3000
Timer1.Enabled = True
End Sub
Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = False
Command1.Font.Size = 8
Command2.Font.Bold = False
Command2.Font.Size = 8

End Sub
