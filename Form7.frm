VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Add_Stock 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00004040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD STOCK DETAILS"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   9734.399
   ScaleMode       =   0  'User
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   600
      Top             =   7800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStock 
      Caption         =   "STOCK REPORT"
      Height          =   375
      Left            =   4800
      Picture         =   "Form7.frx":14A921
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   1695
   End
   Begin VB.ComboBox ComboFuel_Type 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form7.frx":14AF74
      Left            =   5280
      List            =   "Form7.frx":14AF76
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&BACK"
      Height          =   375
      Left            =   6720
      Picture         =   "Form7.frx":14AF78
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&RESET"
      Height          =   375
      Left            =   3480
      Picture         =   "Form7.frx":14B5CB
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&ADD"
      Height          =   375
      Left            =   2040
      Picture         =   "Form7.frx":14BB3B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtCost 
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
      Left            =   5760
      MaxLength       =   9
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
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
      Left            =   5280
      MaxLength       =   5
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3630
      Left            =   10920
      Picture         =   "Form7.frx":14C0AB
      Top             =   6480
      Width           =   4350
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
      Left            =   7080
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   4560
      Width           =   135
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5280
      Picture         =   "Form7.frx":17F93D
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   2160
      X2              =   8280
      Y1              =   5990.4
      Y2              =   5990.4
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   2160
      X2              =   8280
      Y1              =   1747.2
      Y2              =   1747.2
   End
   Begin VB.Label lblCost 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COST PRICE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   8
      Top             =   4680
      Width           =   1290
   End
   Begin VB.Label lblQuantity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   7
      Top             =   3840
      Width           =   1170
   End
   Begin VB.Label lblFuel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FUEL TYPE:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   6
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label lblPlsEnter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   2040
      Width           =   1785
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADD STOCK DETAILS"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   6120
   End
End
Attribute VB_Name = "Add_Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Private rs As New ADODB.Recordset
Private rs1 As New ADODB.Recordset

Private Sub cmdAdd_Click()
If ComboFuel_Type.Text = "" Then
MsgBox "Please Select the Fuel Type"
ComboFuel_Type.SetFocus
Exit Sub

ElseIf txtQuantity.Text = "" Then
MsgBox "Please Enter the Quantity"
txtQuantity.SetFocus
Exit Sub

ElseIf txtCost.Text = "" Then
MsgBox "Please Enter the Cost Price"
txtCost.SetFocus
Exit Sub
End If

rs.AddNew
rs.Fields(0) = ComboFuel_Type.Text
rs.Fields(1) = txtQuantity.Text
rs.Fields(2) = txtCost.Text
rs.Fields(3) = DateTime.Date
rs.Fields(4) = DateTime.Time
rs.Update
rs1.MoveFirst
While Not rs1.EOF
If ComboFuel_Type.Text = rs1.Fields(0) Then
rs1.Fields(1) = rs1.Fields(1) + Val(txtQuantity.Text)
rs1.Update
End If
rs1.MoveNext
Wend
Refresh
txtQuantity.Text = ""
txtCost.Text = ""

End Sub

Private Sub cmdBack_Click()
Unload Me
MainWindow.Show

End Sub

Private Sub cmdReset_Click()
txtQuantity.Text = ""
txtCost.Text = ""

End Sub

Private Sub cmdStock_Click()
Stock_Report.Show

End Sub

Private Sub Form_Load()
ComboFuel_Type.AddItem "Unleaded Petrol"
ComboFuel_Type.AddItem "Premium Petrol"
ComboFuel_Type.AddItem "Diesel"
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DATA.mdb"
rs.Open "select * from AddStock", con, adOpenDynamic, adLockOptimistic
rs1.Open "select * from Stock", con, adOpenDynamic, adLockOptimistic

End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0

End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0

End Sub
