VERSION 5.00
Begin VB.Form frmeditemployee 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form5"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15465
   LinkTopic       =   "Form4"
   ScaleHeight     =   8490
   ScaleWidth      =   15465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "SEARCH"
      Height          =   495
      Left            =   3360
      TabIndex        =   26
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   6600
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   6000
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   3000
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2520
      TabIndex        =   9
      Text            =   "Select Gender"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Qualification"
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
      Left            =   120
      TabIndex        =   23
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
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
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone number"
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
      Left            =   120
      TabIndex        =   21
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Post"
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
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Joining Date"
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
      Left            =   120
      TabIndex        =   19
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Education"
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
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gender"
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
      Left            =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date of Birth"
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
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee Name"
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
      Left            =   -120
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee Id"
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
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit Employee Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   7935
      Left            =   6480
      Picture         =   "editemployee.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmeditemployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New Connection
Dim rs As New Recordset
Dim str As String


Private Sub Command1_Click()
On Error Resume Next
If MsgBox("Data is not recoverable!", vbExclamation + vbOKCancel, "Confirm Delete") = vbOK Then
rs.Delete
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
Text4.Text = ""
Text5.Text = ""
Text9.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text10.Text = ""
Text1.SetFocus
Else
Exit Sub
End If

End Sub


Private Sub Command2_Click()
Unload Me
frmmain.Show
End Sub

Private Sub Command3_Click()
Dim n
Dim str As String
con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\care\Desktop\newproject\mallmanagement.mdb;Persist Security Info=False;"
str = "update employee set Employeename='" & Text2.Text & "', DOB='" & Text3.Text & "', Gender ='" & Combo1.Text & "', Education='" & Text4.Text & "',Salary='" & Text5.Text & "', Qualification='" & Text9.Text & "',address='" & Text6.Text & "',Phonenumber='" & Text7.Text & "',Post='" & Text8.Text & "', joiningdate='" & Text10.Text & "' where Employeeid =" & CInt(Text1.Text) & ""
con.Execute str, n
If n > 0 Then
MsgBox "Record Updated Successfully"
Else
MsgBox "Record Not Found"
End If
con.Close
End Sub

Private Sub Command4_Click()

con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\care\Desktop\newproject\mallmanagement.mdb;Persist Security Info=False;"
rs.Open "select * from employee where Employeeid = " & CInt(Text1.Text) & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "RECORD NOT FOUND", vbInformation, "Search"
Else
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Combo1.Text = rs.Fields(3)
Text4.Text = rs.Fields(4)
Text5.Text = rs.Fields(5)
Text9.Text = rs.Fields(6)
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
Text8.Text = rs.Fields(9)
Text10.Text = rs.Fields(10)

End If
con.Close
End Sub

