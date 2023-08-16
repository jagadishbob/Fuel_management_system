VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00404000&
   BorderStyle     =   0  'None
   Caption         =   "USERS LOGIN FORM"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   MouseIcon       =   "Form2.frx":0000
   Moveable        =   0   'False
   Palette         =   "Form2.frx":08CA
   Picture         =   "Form2.frx":A7C7
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Picture         =   "Form2.frx":12868
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Picture         =   "Form2.frx":12F5C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   5760
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   4680
      Width           =   3975
   End
   Begin VB.Timer timDisplay 
      Interval        =   1000
      Left            =   7560
      Top             =   7440
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 PM"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   9600
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblDay 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   7800
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DAY :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   5
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "OCTOBER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   5280
      TabIndex        =   4
      Top             =   3240
      Width           =   1425
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "2020"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   6420
      TabIndex        =   3
      Top             =   3240
      Width           =   765
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   5010
      TabIndex        =   2
      Top             =   3240
      Width           =   465
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   4200
      Picture         =   "Form2.frx":13650
      Top             =   3120
      Width           =   6945
   End
   Begin VB.Image Image1 
      Height          =   4335
      Index           =   0
      Left            =   4200
      Picture         =   "Form2.frx":1F4E2
      Top             =   3840
      Width           =   6855
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Private rs As New ADODB.Recordset
Public us As String
Public ld As String
Public lt As String

Private Sub cmdExit_Click()
End

End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.Font.Bold = False
cmdExit.Font.Size = 17

End Sub

Private Sub cmdLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdLogin.Font.Bold = False
cmdLogin.Font.Size = 17

End Sub
Private Sub cmdLogin_Click()
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Now
rs.Update
us = Text1.Text
ld = DateTime.Date
lt = DateTime.Time
Refresh
If Text1.Text = "jagadish" And Text2.Text = "123" Then
Loading.Show
Unload Login
ElseIf Text1.Text = "saiteja" And Text2.Text = "124" Then
Loading.Show
Unload Login
ElseIf Text1.Text = "harshith" And Text2.Text = "125" Then
Loading.Show
Unload Login
ElseIf Text1.Text = "" Then
MsgBox "Please Enter the USERNAME"
Text1.SetFocus
Exit Sub
ElseIf Text2.Text = "" Then
MsgBox "Please Enter the Appropriate PASSWORD"
Text2.SetFocus
Exit Sub
Else
Text1.Text = ""
Text2.Text = ""
MsgBox "The User Failed to Login"
Text1.SetFocus
End If

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DATA.mdb"
rs.Open "select * from Authentication", con, adOpenDynamic, adLockOptimistic
'rs1.Open "select * from LoginInfo", con, adOpenDynamic, adLockOptimistic

Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdLogin.Font.Bold = False
cmdLogin.Font.Size = 16
cmdExit.Font.Bold = False
cmdExit.Font.Size = 16

End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub

Private Sub timDisplay_Timer()
Dim Today As Variant
Today = Now
lblDay.Caption = Format(Today, "dddd")
lblMonth.Caption = Format(Today, "mmmm")
lblYear.Caption = Format(Today, "yyyy")
lblNumber.Caption = Format(Today, "dd")
lblTime.Caption = Format(Today, "h:mm:ss ampm")
End Sub
