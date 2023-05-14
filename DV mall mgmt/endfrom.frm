VERSION 5.00
Begin VB.Form frmreport 
   Caption         =   "Form13"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9975
   LinkTopic       =   "Form13"
   ScaleHeight     =   7335
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   " Shop Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Tenent Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Employee Attendance"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Rent Collection"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   0
      Picture         =   "endfrom.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


























































































































































































Private Sub Command2_Click()
frmshopreport.Show
frmreport.Hide
End Sub

Private Sub Command3_Click()
frmemployeereport.Show
frmreport.Hide
End Sub

Private Sub Command4_Click()
frmtenentreport.Show
frmreport.Hide
End Sub

Private Sub Command6_Click()
frmrentreport.Show
End Sub

Private Sub Command7_Click()
frmattendancereport.Show
frmreport.Hide
End Sub

Private Sub Command9_Click()
Unload Me
frmmain.Show
End Sub
