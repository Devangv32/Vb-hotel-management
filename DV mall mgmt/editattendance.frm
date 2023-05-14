VERSION 5.00
Begin VB.Form frmeditattendance 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form12"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11955
   LinkTopic       =   "Form12"
   ScaleHeight     =   5625
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "SEARCH"
      Height          =   615
      Left            =   2760
      TabIndex        =   15
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "editattendance.frx":0000
      Left            =   3360
      List            =   "editattendance.frx":0022
      TabIndex        =   6
      Top             =   840
      Width           =   3015
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "editattendance.frx":0045
      Left            =   3360
      List            =   "editattendance.frx":004F
      TabIndex        =   5
      Text            =   "Select Evening Status"
      Top             =   4080
      Width           =   3015
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "editattendance.frx":0064
      Left            =   3360
      List            =   "editattendance.frx":006E
      TabIndex        =   4
      Text            =   "Select Morning Status"
      Top             =   3432
      Width           =   3015
   End
   Begin VB.ComboBox Combo4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "editattendance.frx":0083
      Left            =   3360
      List            =   "editattendance.frx":00AB
      TabIndex        =   3
      Top             =   2784
      Width           =   3015
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "editattendance.frx":00F7
      Left            =   3360
      List            =   "editattendance.frx":011F
      TabIndex        =   2
      Top             =   2136
      Width           =   3015
   End
   Begin VB.ComboBox Combo6 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "editattendance.frx":0169
      Left            =   3360
      List            =   "editattendance.frx":01C4
      TabIndex        =   1
      Top             =   1488
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   6600
      Picture         =   "editattendance.frx":0234
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "EDIT EMPLOYEE ATTENDANCE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   13
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee Id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   800
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   1456
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   2112
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   2768
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Morning Status"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   3424
      Width           =   3015
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Evening Status"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   3015
   End
End
Attribute VB_Name = "frmeditattendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New Connection
Dim rs As New Recordset


Private Sub Command2_Click()
Unload Me
frmmain.Show
End Sub

Private Sub Command3_Click()
Dim n
Dim str As String
con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\care\Desktop\newproject\mallmanagement.mdb;Persist Security Info=False"
str = "update Employeeattendance set Datee= '" & Combo6.Text & "', Monthh = '" & Combo5.Text & "', Yearr='" & Combo4.Text & "', Morningstatus= '" & Combo3.Text & "', Eveningstatus='" & Combo2.Text & "' where Employeeid =" & CInt(Combo1.Text) & ""
'str = "update rent set Receiptnumber='" & Combo1.Text & "', tenentname ='" & Text4.Text & "', rentamount ='" & Text3.Text & "', rentformonth='" & Combo3.Text & "',rentforyear='" & Combo2.Text & "', rentpaiddate ='" & Text2.Text & "' where shopid =" & CInt(Text1.Text) & ""
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
rs.Open "select * from Employeeattendance where Employeeid = " & CInt(Combo1.Text) & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "RECORD NOT FOUND", vbInformation, "Search"
Else
Combo1.Text = rs.Fields(0)
Combo6.Text = rs.Fields(1)
Combo5.Text = rs.Fields(2)
Combo4.Text = rs.Fields(3)
Combo3.Text = rs.Fields(4)
Combo2.Text = rs.Fields(5)

End If
con.Close
End Sub


