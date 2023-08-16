VERSION 5.00
Begin VB.Form Change_Rates 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "FUEL RATES & CHANGING THE FUEL RATES"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   FillColor       =   &H0080FF80&
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form14.frx":0000
   ScaleHeight     =   7755
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&UPDATE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Picture         =   "Form14.frx":36389
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF8080&
      Caption         =   "&B A C K"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      MaskColor       =   &H00FFC0C0&
      Picture         =   "Form14.frx":3698E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmdChange_Rates 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&CHANGE RATES"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      Picture         =   "Form14.frx":36F93
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox txtDiesel 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtPremium 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtUnleaded 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   4530
      Left            =   -600
      Picture         =   "Form14.frx":37598
      Top             =   2040
      Width           =   2790
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   2520
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   2
      Left            =   7320
      Picture         =   "Form14.frx":60A7A
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   1
      Left            =   7320
      Picture         =   "Form14.frx":61590
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   7320
      Picture         =   "Form14.frx":620A6
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   3120
      X2              =   7800
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unleaded petrol"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Premium petrol"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Diesel"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FUEL RATES"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "Change_Rates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Private rs As New ADODB.Recordset
Dim X As Integer

Private Sub cmdBack_Click()
Unload Me
MainWindow.Show

End Sub

Private Sub cmdChange_Rates_Click()
If X = 0 Then
Manager.Show
X = X + 1
End If
txtUnleaded.Enabled = True
txtPremium.Enabled = True
txtDiesel.Enabled = True

End Sub

Private Sub cmdUpdate_Click()
rs.Fields(0) = txtUnleaded.Text
rs.Fields(1) = txtPremium.Text
rs.Fields(2) = txtDiesel.Text
rs.Update
MsgBox ("Fuel Rates Changed")

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DATA.mdb"
rs.Open "select * from Rates", con, adOpenDynamic, adLockOptimistic

txtUnleaded.Text = rs.Fields(0)
txtPremium.Text = rs.Fields(1)
txtDiesel.Text = rs.Fields(2)

End Sub


Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub

Private Sub txtDiesel_KeyPress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 32

End Sub

Private Sub txtPremium_KeyPress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 32

End Sub

Private Sub txtUnleaded_KeyPress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 32

End Sub
