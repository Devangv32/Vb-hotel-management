VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   11
      Top             =   2880
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
      TabIndex        =   10
      Top             =   3840
      Width           =   3375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Reports"
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
      TabIndex        =   9
      Top             =   4920
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
      TabIndex        =   8
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Edit Employee Attendance"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Edit Tenent details "
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
      TabIndex        =   6
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Edit Rent Collection"
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
      Top             =   3360
      Width           =   3375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Edit Employee Details"
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
      Top             =   1440
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
      Top             =   1920
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
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit Shop"
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
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "add shop"
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
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   7215
      Left            =   0
      Picture         =   "shopdeatil.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmshop.Show
frmmain.Hide
End Sub

Private Sub Command10_Click()
frmeditattendance.Show
frmmain.Hide
End Sub

Private Sub Command11_Click()
frmeditemployee.Show
frmmain.Hide
End Sub

Private Sub Command12_Click()
frmedittenent.Show
frmmain.Hide
End Sub

Private Sub Command13_Click()
frmeditrent.Show
frmmain.Hide
End Sub

Private Sub Command2_Click()
frmshop.Show

End Sub

Private Sub Command3_Click()
frmemployee.Show
frmmain.Hide
End Sub

Private Sub Command4_Click()
frmtenent.Show
frmmain.Hide
End Sub

Private Sub Command6_Click()
frmrent.Show
frmmain.Hide
End Sub

Private Sub Command7_Click()
frmemployeeattendance.Show
frmmain.Hide
End Sub

Private Sub Command8_Click()
frmreport.Show
frmmain.Hide
End Sub

Private Sub Command9_Click()
Unload Me
frmlogin.Show
End Sub
