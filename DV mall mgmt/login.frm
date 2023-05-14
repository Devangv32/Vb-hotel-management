VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "Login"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9810
   LinkTopic       =   "Form14"
   ScaleHeight     =   6600
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
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
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "password"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "username"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   6975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   -360
      Width           =   9855
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String



Private Sub Command1_Click()
Set rs = New ADODB.Recordset
str = "select * from login"
str = str & " where username='" & Text1.Text & "' and "
str = str & " password = '" & Text2.Text & "'"
rs.Open str, con, adOpenKeyset, adLockOptimistic
If rs.EOF Then
MsgBox "Invalid login details", vbExclamation, " login incorrect "
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Text2.SetFocus
Exit Sub
End If
rs.Close
Set rs = Nothing
frmmain.Show
End Sub

Private Sub Command2_Click()

Unload Me
End Sub


Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\care\Desktop\newproject\mallmanagement.mdb;Persist Security Info=False"
con.Open
End Sub
