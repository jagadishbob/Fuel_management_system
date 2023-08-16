VERSION 5.00
Begin VB.Form Search 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEARCHING"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   9360
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboVehicleNo 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtBillNo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFFF&
      Caption         =   "BACK TO MAIN WINDOW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Picture         =   "Form11.frx":3A869
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8760
      Width           =   3015
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "S E A R C H"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      Picture         =   "Form11.frx":3B066
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Search"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtQuantity 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox txtProduct 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtCustomerName 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txtDate 
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox txtTime 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   10920
      Picture         =   "Form11.frx":3B8C8
      Top             =   6960
      Width           =   4365
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
      Left            =   8640
      TabIndex        =   19
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   7320
      TabIndex        =   18
      Top             =   6600
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   6960
      Picture         =   "Form11.frx":680CA
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   405
   End
   Begin VB.Label lblBillNo 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL NUMBER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE SELECT THE VEHICLE NUMBER:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2280
      TabIndex        =   15
      Top             =   1920
      Width           =   4440
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATE PURCHASED"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   14
      Top             =   7320
      Width           =   2520
   End
   Begin VB.Label lblQuantity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY PURCHASED"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   13
      Top             =   5880
      Width           =   2745
   End
   Begin VB.Label lblProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT PURCHASED"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   12
      Top             =   5160
      Width           =   2670
   End
   Begin VB.Label lblCustomerName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   11
      Top             =   4440
      Width           =   2160
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIME PURCHASED"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   10
      Top             =   8040
      Width           =   2145
   End
   Begin VB.Label lblAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL BILL AMOUNT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Top             =   6600
      Width           =   2550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPECIFIC  SEARCH"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   4200
      TabIndex        =   0
      Top             =   600
      Width           =   3885
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Private rs As New ADODB.Recordset
Private rs1 As New ADODB.Recordset
Private Sub cmdSearch_Click()
Dim N As Integer
Dim X As Integer
N = 0
txtBillNo.Text = " "
txtCustomerName.Text = " "
txtProduct.Text = " "
txtQuantity.Text = " "
txtAmount.Text = " "
txtDate.Text = " "
txtTime.Text = " "
rs.MoveFirst
If ComboVehicleNo.Text = "" Then
MsgBox "PLEASE SELECT VEHICLE NUMBER"
ComboVehicleNo.SetFocus
End If
While Not rs.EOF
If ComboVehicleNo.Text = rs.Fields(2) Then
txtBillNo.Text = rs.Fields(0)
txtCustomerName.Text = rs.Fields(1)
txtProduct.Text = rs.Fields(3)
txtQuantity.Text = rs.Fields(4)
txtAmount.Text = rs.Fields(5)
txtDate.Text = rs.Fields(6)
txtTime.Text = rs.Fields(7)
N = 1
If X = vbNo Then
Exit Sub
End If
End If
rs.MoveNext
Wend
If N = 0 Then
MsgBox ("Search unsuccessful, record not found")
Else
MsgBox ("Search completed Successfully")
End If
End Sub

Sub showrec()

End Sub

Private Sub cmdSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSearch.Font.Size = 9
cmdSearch.Font.Bold = True

End Sub

Private Sub cmdBack_Click()
Unload Me
MainWindow.Show
End Sub

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdBack.Font.Size = 9
cmdBack.Font.Bold = True
End Sub
Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DATA.mdb"
rs.Open "select * from DataEntry", con, adOpenDynamic, adLockOptimistic
rs.MoveFirst
While Not rs.EOF
ComboVehicleNo.AddItem (rs(2))
rs.MoveNext
Wend
cmdSearch.Enabled = True

End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSearch.Font.Size = 8
cmdSearch.Font.Bold = False
cmdBack.Font.Size = 8
cmdBack.Font.Bold = False


End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub
