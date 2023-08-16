VERSION 5.00
Begin VB.Form View_Stock 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIEW STOCK DETAILS"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   10140
   ScaleMode       =   0  'User
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "BACK TO &MAIN WINDOW"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      Picture         =   "Form8.frx":3BC72
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton cmdDataEntry 
      Caption         =   "BACK TO &DATA ENTRY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      Picture         =   "Form8.frx":3C277
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddStock 
      Caption         =   "BACK TO &ADDSTOCK"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Picture         =   "Form8.frx":3C87C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox txtDiesel 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtPremium 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtUnleaded 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   3000
      X2              =   7680
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Image Image1 
      Height          =   3780
      Left            =   11280
      Picture         =   "Form8.frx":3CE81
      Top             =   6360
      Width           =   4005
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Litres "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   8160
      TabIndex        =   12
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Litres "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   8160
      TabIndex        =   11
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Litres "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   10
      Top             =   4200
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   1920
      X2              =   9000
      Y1              =   2470
      Y2              =   2470
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   1800
      X2              =   9240
      Y1              =   7020
      Y2              =   7020
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DIESEL"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PREMIUM PETROL"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "UNLEADED PETROL"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VIEW STOCK DETAILS"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   3000
      TabIndex        =   0
      Top             =   1320
      Width           =   4665
   End
End
Attribute VB_Name = "View_Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Private rs As New ADODB.Recordset

Private Sub cmdAddStock_Click()
Add_Stock.Show

End Sub

Private Sub cmdBack_Click()
Unload Me
MainWindow.Show

End Sub

Private Sub cmdDataEntry_Click()
DataEntry.Show

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DATA.mdb"
rs.Open "select * from Stock", con, adOpenDynamic, adLockOptimistic
If rs.Fields(0) = "Unleaded Petrol" Then
txtUnleaded.Text = rs.Fields(1)
End If
rs.MoveNext
If rs.Fields(0) = "Premium Petrol" Then
txtPremium.Text = rs.Fields(1)
End If
rs.MoveNext
If rs.Fields(0) = "Diesel" Then
txtDiesel.Text = rs.Fields(1)
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub
