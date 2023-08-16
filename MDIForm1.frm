VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainWindow 
   BackColor       =   &H8000000C&
   Caption         =   "FUEL AUTOMATION SYSTEM MAIN WINDOW"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   355
      Top             =   9840
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   20190
      TabIndex        =   0
      Top             =   10260
      Width           =   20250
      Begin VB.Timer DisplayTime 
         Index           =   1
         Interval        =   1000
         Left            =   10680
         Top             =   0
      End
      Begin VB.Label lblTim 
         Caption         =   "00:00:00 AM"
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
         Left            =   14040
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblYear 
         Caption         =   "2011"
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
         Left            =   13440
         TabIndex        =   5
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblMonth 
         Caption         =   "September"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12360
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblNumber 
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12000
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblDay 
         Caption         =   "Thursday"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   10560
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " Fuel Automation System. @  Platform VB 6.0"
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
         Left            =   -240
         TabIndex        =   1
         Top             =   0
         Width           =   7455
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   -240
      Top             =   9120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":21206A
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2120C8
            Key             =   "New"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":212126
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":212184
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2121E2
            Key             =   "Back"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCustomers 
      Caption         =   "&CUSTOMERS"
      Begin VB.Menu mnuDataEntry 
         Caption         =   "DATA ENTRY"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuCustomer_Details 
         Caption         =   "&CUSTOMER DETAILS"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&SEARCH"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "&USERS"
      Begin VB.Menu mnuUser_Details 
         Caption         =   "&USER DETAILS"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLogin_Details 
         Caption         =   "&USERS LOGIN DETAILS"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuChange_Rates 
         Caption         =   "&CHANGE FUEL RATES"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuStock 
      Caption         =   "&STOCK"
      Begin VB.Menu mnuAdd_Stock 
         Caption         =   "&ADD_STOCK"
      End
      Begin VB.Menu mnuStock_Details 
         Caption         =   "&STOCK DETAILS"
      End
      Begin VB.Menu mnuView_Stock 
         Caption         =   "&VIEW_STOCK"
      End
   End
   Begin VB.Menu mnuTransactions 
      Caption         =   "&TRANSACTIONS"
      Begin VB.Menu mnuDateWiseTrans 
         Caption         =   "&DATEWISE TRANSACTIONS"
      End
      Begin VB.Menu mnuFuelWiseTrans 
         Caption         =   "&FUEL WISE DAILY TRANSACTIONS"
      End
      Begin VB.Menu mnuDateWiseStock 
         Caption         =   "&DATEWISE FUEL STOCK "
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&REPORTS"
      Begin VB.Menu mnuStock_Report 
         Caption         =   "&STOCK REPORT"
      End
      Begin VB.Menu mnuSales_Report 
         Caption         =   "&SALES REPORT"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&WINDOW"
      Begin VB.Menu mnuLogout 
         Caption         =   "&LOGOUT"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&CLOSE"
         Shortcut        =   +{DEL}
      End
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private con As New ADODB.Connection
Private rs As New ADODB.Recordset
Private Sub DisplayTime_Timer(Index As Integer)
Dim Today As Variant
Today = Now
lblDay.Caption = Format(Today, "dddd")
lblNumber.Caption = Format(Today, "dd")
lblMonth.Caption = Format(Today, "mmmm")
lblYear.Caption = Format(Today, "yyyy")
lblTim.Caption = Format(Today, "h:mm:ss ampm")

End Sub

Private Sub MDIForm_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DATA.mdb"
rs.Open "select * from LoginInfo", con, adOpenDynamic, adLockOptimistic

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
con.Close

End Sub

Private Sub mnuAdd_Stock_Click()
Add_Stock.Show

End Sub

Private Sub mnuChange_Rates_Click()
Change_Rates.Show

End Sub

Private Sub mnuClose_Click()
rs.AddNew
rs.Fields(0) = Login.us
rs.Fields(1) = Login.ld
rs.Fields(2) = Login.lt
rs.Fields(3) = Date
rs.Fields(4) = Time
rs.Update

End

End Sub

Private Sub mnuCustomer_Details_Click()
Customer_Details.Show

End Sub

Private Sub mnuDataEntry_Click()
DataEntry.Show

End Sub

Private Sub mnuDateWiseStock_Click()
DateWiseStock.Show

End Sub

Private Sub mnuDateWiseTrans_Click()
DateWiseTrans.Show

End Sub

Private Sub mnuFuelWiseTrans_Click()
DateWiseFuelTrans.Show

End Sub

Private Sub mnuLogin_Details_Click()
Login_Details.Show

End Sub

Private Sub mnuLogout_Click()
rs.AddNew
rs.Fields(0) = Login.us
rs.Fields(1) = Login.ld
rs.Fields(2) = Login.lt
rs.Fields(3) = Date
rs.Fields(4) = Time
rs.Update
Login.Show
Unload MainWindow

End Sub

Private Sub mnuSales_Report_Click()
Sales_Report.Show

End Sub

Private Sub mnuSearch_Click()
Search.Show

End Sub

Private Sub mnuStock_Details_Click()
Stock_Details.Show

End Sub

Private Sub mnuStock_Report_Click()
Stock_Report.Show

End Sub

Private Sub mnuUser_Details_Click()
User_Details.Show

End Sub

Private Sub mnuView_Stock_Click()
View_Stock.Show

End Sub

Private Sub Timer1_Timer()
If Label1.Left <= 5300 Then
Label1.Left = Label1.Left + 200
Else
Label1.Left = -5300
End If

End Sub
