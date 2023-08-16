VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form DataEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMERS DATA ENTRY "
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MouseIcon       =   "Form6.frx":1EAF2
   Moveable        =   0   'False
   Picture         =   "Form6.frx":2013C
   ScaleHeight     =   9360
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   11160
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
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
      Connect         =   $"Form6.frx":5DF42
      OLEDBString     =   $"Form6.frx":5DFE1
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&BACK TO MAIN WINDOW"
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
      Left            =   8040
      Picture         =   "Form6.frx":5E080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton cmdStock 
      Caption         =   "&FUEL STOCK"
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
      Picture         =   "Form6.frx":5E685
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdRates 
      Caption         =   "&FUEL RATES"
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
      Left            =   4920
      Picture         =   "Form6.frx":5EC38
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew_Reset 
      Caption         =   "&NEW / RESET"
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
      Left            =   3240
      Picture         =   "Form6.frx":5F1EB
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE RECORD"
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
      Left            =   1560
      Picture         =   "Form6.frx":5F79E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&CALCULATE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Picture         =   "Form6.frx":5FD51
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   5520
      Width           =   1095
   End
   Begin VB.ComboBox ComProduct 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4800
      Width           =   2535
   End
   Begin VB.TextBox txtVehicleNo 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtCustomerName 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox txtBillNo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   10800
      Picture         =   "Form6.frx":60304
      Top             =   7320
      Width           =   4470
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
      Left            =   6120
      TabIndex        =   19
      Top             =   5640
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   1560
      X2              =   10680
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      Picture         =   "Form6.frx":88AC6
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      X1              =   2040
      X2              =   7440
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblBillNo 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL NUMBER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA ENTRY FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE ENTER ALL THE DETAILS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   5295
   End
   Begin VB.Label lblCust_Name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   2040
      TabIndex        =   3
      Top             =   3360
      Width           =   2100
   End
   Begin VB.Label lblProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   2040
      TabIndex        =   2
      Top             =   4800
      Width           =   1140
   End
   Begin VB.Label lblQuantity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblVehicleNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VEHICLE NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   4080
      Width           =   1410
   End
End
Attribute VB_Name = "DataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Private rs As New ADODB.Recordset
Private rs1 As New ADODB.Recordset
Private rs2 As New ADODB.Recordset
Dim num As Integer

Private Sub cmdBack_Click()
Unload Me
MainWindow.Show

End Sub
Private Sub cmdCalc_Click()
If txtCustomerName.Text = "" Then
MsgBox "Please Enter the Customer Name"
txtCustomerName.SetFocus
Exit Sub

ElseIf txtVehicleNo.Text = "" Then
MsgBox "Please Enter Vehicle Number"
txtVehicleNo.SetFocus
Exit Sub

ElseIf txtQuantity.Text = "" Then
MsgBox "Please Enter the Quantity (in Litres)"
txtQuantity.SetFocus
Exit Sub
End If

rs2.MoveFirst
If rs2.Fields(1) <= 10 Then
MsgBox ("Unleaded Petrol out of Stock")
End If
rs2.MoveNext
If rs2.Fields(1) <= 10 Then
MsgBox ("Premium Petrol Out Of Stock")
End If
rs2.MoveNext
If rs2.Fields(1) <= 10 Then
MsgBox ("Diesel out of Stock")
End If
If ComProduct.Text = "Unleaded Petrol" Then
txtAmount.Text = (txtQuantity.Text * rs1.Fields(0))
End If
If ComProduct.Text = "Premium Petrol" Then
txtAmount.Text = (txtQuantity.Text * rs1.Fields(1))
End If
If ComProduct.Text = "Diesel" Then
txtAmount.Text = (txtQuantity.Text * rs1.Fields(2))
End If

End Sub
Private Sub cmdNew_Reset_Click()
Call AutoReg
txtBillNo = Format(num, "B000")
txtCustomerName.Text = ""
txtVehicleNo.Text = ""
txtQuantity.Text = ""
txtAmount.Text = ""
txtCustomerName.SetFocus

End Sub
Private Sub cmdRates_Click()
Rates.Show

End Sub
Private Sub cmdSave_Click()
If txtCustomerName.Text = "" Then
MsgBox "Please Enter the Customer Name"
txtCustomerName.SetFocus
Exit Sub

ElseIf txtVehicleNo.Text = "" Then
MsgBox "Please Enter Vehicle Number"
txtVehicleNo.SetFocus
Exit Sub

ElseIf txtQuantity.Text = "" Then
MsgBox "Please Enter the Quantity (in Litres)"
txtQuantity.SetFocus
Exit Sub

ElseIf txtAmount.Text = "" Then
MsgBox "Please Click the Calculate Button"
cmdCalc.SetFocus
Exit Sub
End If

rs2.MoveFirst
If rs2.Fields(1) <= 10 Then
MsgBox ("Unleaded Petrol out of Stock")
End If
rs2.MoveNext
If rs2.Fields(1) <= 10 Then
MsgBox ("Premium Petrol Out Of Stock")
End If
rs2.MoveNext
If rs2.Fields(1) <= 10 Then
MsgBox ("Diesel out of Stock")
End If

rs.AddNew
rs.Fields(0) = txtBillNo.Text
rs.Fields(1) = txtCustomerName.Text
rs.Fields(2) = txtVehicleNo.Text
rs.Fields(3) = ComProduct.Text
rs.Fields(4) = txtQuantity.Text
rs.Fields(5) = txtAmount.Text
rs.Fields(6) = DateTime.Date
rs.Fields(7) = DateTime.Time
rs.Update
MsgBox "Record Updated."

rs2.MoveFirst
While Not rs2.EOF
If ComProduct.Text = rs2.Fields(0) Then
rs2.Fields(1) = rs2.Fields(1) - Val(txtQuantity.Text)
rs2.Update
End If
rs2.MoveNext
Wend
Refresh
Call AutoReg
txtBillNo.Text = ""
txtCustomerName.Text = ""
txtVehicleNo.Text = ""
'ComProduct.Text = ""
txtQuantity.Text = ""
txtAmount.Text = ""

End Sub
Private Sub cmdStock_Click()
Stock_Report.Show

End Sub
Private Sub Form_Load()
Me.Show
ComProduct.AddItem "Unleaded Petrol"
ComProduct.AddItem "Premium Petrol"
ComProduct.AddItem "Diesel"
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DATA.mdb"
rs.Open "select * from DataEntry", con, adOpenDynamic, adLockOptimistic
rs1.Open "select * from Rates", con, adOpenDynamic, adLockOptimistic
rs2.Open "select * from Stock", con, adOpenDynamic, adLockOptimistic
Call AutoReg
txtBillNo = Format(num, "B000")
txtCustomerName.Text = ""
txtVehicleNo.Text = ""
txtQuantity.Text = ""
txtAmount.Text = ""
txtCustomerName.SetFocus

End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub
Public Sub AutoReg()
On Error GoTo ErrMsg
Refresh
If rs.RecordCount = 0 Then
num = 1
Else
rs.MoveLast
num = Mid(rs(0), 2, 3) + 1
End If
Exit Sub
ErrMsg:
MsgBox "No Records Found"

End Sub


Private Sub txtCustomerName_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then KeyAscii = 0

End Sub


Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0

End Sub
