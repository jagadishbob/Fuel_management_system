VERSION 5.00
Begin VB.Form Customer_Details 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMERS DETAILS"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   9360
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBillNo 
      DataField       =   "Billno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtCustomerName 
      DataField       =   "Cust Name"
      DataSource      =   "Adodc1"
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
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtVehicleNo 
      DataField       =   "Vehi No"
      DataSource      =   "Adodc1"
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
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox txtProduct 
      DataField       =   "Pro"
      DataSource      =   "Adodc1"
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
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      DataField       =   "Quantity"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txtDate 
      DataField       =   "curr_date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox txtTime 
      DataField       =   "Time"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox txtAmount 
      DataField       =   "Total"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Picture         =   "Form9.frx":3E096
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Picture         =   "Form9.frx":3E893
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Prev Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Picture         =   "Form9.frx":3F090
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1200
      Picture         =   "Form9.frx":3F88D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back to MainWindow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7920
      Picture         =   "Form9.frx":4008A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   3270
      Left            =   11040
      Picture         =   "Form9.frx":40887
      Top             =   6720
      Width           =   4305
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
      TabIndex        =   23
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6120
      TabIndex        =   22
      Top             =   6720
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2760
      TabIndex        =   21
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL NUMBER"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER NAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "VEHICLE NUMBER"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT PURCHASED"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   17
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY PURCHASED"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL AMOUNT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF PURCHASED"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME OF PURCHASED"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   2
      X1              =   1080
      X2              =   10320
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5760
      Picture         =   "Form9.frx":6E889
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   405
   End
End
Attribute VB_Name = "Customer_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Private rs As New ADODB.Recordset

Private Sub cmdBack_Click(Index As Integer)
Unload Me
MainWindow.Show

End Sub

Private Sub cmdFirst_Click(Index As Integer)
rs.MoveFirst
txtBillNo.Text = rs.Fields(0)
txtCustomerName.Text = rs.Fields(1)
txtVehicleNo.Text = rs.Fields(2)
txtProduct.Text = rs.Fields(3)
txtQuantity.Text = rs.Fields(4)
txtAmount.Text = rs.Fields(5)
txtDate.Text = rs.Fields(6)
txtTime.Text = rs.Fields(7)

End Sub

Private Sub cmdLast_Click()
rs.MoveLast
txtBillNo.Text = rs.Fields(0)
txtCustomerName.Text = rs.Fields(1)
txtVehicleNo.Text = rs.Fields(2)
txtProduct.Text = rs.Fields(3)
txtQuantity.Text = rs.Fields(4)
txtAmount.Text = rs.Fields(5)
txtDate.Text = rs.Fields(6)
txtTime.Text = rs.Fields(7)

End Sub

Private Sub cmdNext_Click()
If rs.EOF Then
rs.MoveFirst
End If
If rs.BOF Then
rs.MoveLast
End If
txtBillNo.Text = rs.Fields(0)
txtCustomerName.Text = rs.Fields(1)
txtVehicleNo.Text = rs.Fields(2)
txtProduct.Text = rs.Fields(3)
txtQuantity.Text = rs.Fields(4)
txtAmount.Text = rs.Fields(5)
txtDate.Text = rs.Fields(6)
txtTime.Text = rs.Fields(7)
rs.MoveNext

End Sub

Private Sub cmdPrevious_Click()
If rs.BOF Then
rs.MoveLast
End If
If rs.EOF Then
rs.MoveFirst
End If
txtBillNo.Text = rs.Fields(0)
txtCustomerName.Text = rs.Fields(1)
txtVehicleNo.Text = rs.Fields(2)
txtProduct.Text = rs.Fields(3)
txtQuantity.Text = rs.Fields(4)
txtAmount.Text = rs.Fields(5)
txtDate.Text = rs.Fields(6)
txtTime.Text = rs.Fields(7)
rs.MovePrevious

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DATA.mdb"
rs.Open "select * from DataEntry", con, adOpenDynamic, adLockOptimistic

txtBillNo.Text = rs.Fields(0)
txtCustomerName.Text = rs.Fields(1)
txtVehicleNo.Text = rs.Fields(2)
txtProduct.Text = rs.Fields(3)
txtQuantity.Text = rs.Fields(4)
txtAmount.Text = rs.Fields(5)
txtDate.Text = rs.Fields(6)
txtTime.Text = rs.Fields(7)

End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub

