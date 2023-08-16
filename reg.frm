VERSION 5.00
Begin VB.Form Starting 
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
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   960
      Top             =   5640
   End
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
      Left            =   3120
      ScaleHeight     =   615
      ScaleWidth      =   7800
      TabIndex        =   2
      Top             =   1200
      Width           =   7800
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "&CANCEL"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&OK"
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   3345
      Left            =   4800
      Top             =   2160
      Width           =   4500
   End
End
Attribute VB_Name = "Starting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
