VERSION 5.00
Begin VB.Form Starting 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FUEL AUTOMATION SYSTEM STARTING......"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   FillColor       =   &H00808080&
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4890
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&CANCEL"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Picture         =   "Form1.frx":22724
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Picture         =   "Form1.frx":41216
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel Automation System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   7815
   End
   Begin VB.Image Image1 
      Height          =   2190
      Left            =   3480
      Picture         =   "Form1.frx":483B0
      Top             =   1440
      Width           =   3045
   End
End
Attribute VB_Name = "Starting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
End

End Sub
Private Sub cmdOK_Click()
Login.Show
Unload Me

End Sub

