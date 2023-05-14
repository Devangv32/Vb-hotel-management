VERSION 5.00
Begin VB.Form frmshop 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15585
   LinkTopic       =   "Form2"
   ScaleHeight     =   8010
   ScaleWidth      =   15585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      DataField       =   "Bonus Amount"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   6840
      Width           =   4335
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "Status"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "shopdeatil2.frx":0000
      Left            =   2640
      List            =   "shopdeatil2.frx":000A
      TabIndex        =   12
      Top             =   6240
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Floor number"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Top             =   4560
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdsave_click_err 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "SAVE"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "Rent Per Month"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Top             =   5760
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      DataField       =   "Shop Number"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   5
      Top             =   5160
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      DataField       =   "Shop id"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bonus Amount"
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
      TabIndex        =   13
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status"
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
      Left            =   0
      TabIndex        =   11
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rent Per Month"
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
      TabIndex        =   10
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shop Number"
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
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " Floor Number"
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
      TabIndex        =   2
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shop Id"
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
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shop Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   4695
      Left            =   17760
      Picture         =   "shopdeatil2.frx":0027
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   0
      Picture         =   "shopdeatil2.frx":3FA1CF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15495
   End
End
Attribute VB_Name = "frmshop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New Connection
Dim rs As New Recordset
Dim str As String

Private Sub cmdsave_click_err_Click()
rs.Open "insert into shop values (" & Text1.Text & ",'" & Combo1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Combo2.Text & "','" & Text4.Text & "' )", con, adOpenDynamic, adLockOptimistic
MsgBox ("Shop Details got successfully added.")
ClearAllFields
End Sub
Private Sub ClearAllFields()
Text1.Text = ""
Combo1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo2.Text = ""
Text4.Text = ""

End Sub
Private Sub Command2_Click()
Unload Me
frmmain.Show
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\care\Desktop\newproject\mallmanagement.mdb;Persist Security Info=False"
con.Open
'MsgBox ("connected to datababase")


End Sub
