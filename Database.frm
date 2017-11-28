VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form18 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form18"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   LinkTopic       =   "Form18"
   ScaleHeight     =   6465
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   5880
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Database.frx":0000
      OLEDBString     =   $"Database.frx":008E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Emp_Details"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   240
      TabIndex        =   33
      Top             =   3120
      Width           =   4575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Edit to Database"
      Height          =   495
      Left            =   2760
      TabIndex        =   30
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Remove Data"
      Height          =   495
      Left            =   240
      TabIndex        =   29
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save to Database"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   28
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add Data"
      Height          =   495
      Left            =   240
      TabIndex        =   27
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Last"
      Height          =   495
      Left            =   8280
      TabIndex        =   26
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "First"
      Height          =   495
      Left            =   6240
      TabIndex        =   25
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next"
      Height          =   495
      Left            =   2160
      TabIndex        =   24
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Previous"
      Height          =   495
      Left            =   4200
      TabIndex        =   23
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   21
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   20
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   19
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   14
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   13
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   12
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "Emp_Basic"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "Emp_Desig"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "Emp_Dept"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "Emp_Name"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Extra Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   32
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Basic Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Net Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Income Tax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gross Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CCA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "HRA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "DA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Basic Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sname As String
sname = InputBox("Enter the Name You want to Search: ")
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("Emp_Name = '" & sname & "'")
If Text1.Text = "" Then
    MsgBox ("No Data Found")
    Adodc1.Recordset.MoveFirst
End If
Command10_Click
End Sub

Private Sub Command10_Click()
Text5.Text = Val(Text4.Text) * 0.52
Text6.Text = Val(Text4.Text) * 0.16
Text7.Text = Val(Text4.Text) * 0.08
Text8.Text = Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text)
Text9.Text = Val(Text4.Text) * 0.06
If Val(Text4.Text) >= 80000 Then
    Text10.Text = Val(Text4.Text) * 0.08
Else
    Text10.Text = 0
End If
Text11.Text = Val(Text8.Text) - Val(Text9.Text) - Val(Text10.Text)
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious
Command10_Click
If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveFirst
    Command10_Click
End If
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext
Command10_Click
If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveLast
    Command10_Click
End If
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveFirst
Command10_Click
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveLast
Command10_Click
End Sub

Private Sub Command6_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Command7.Enabled = True
Adodc1.Recordset.AddNew
Command10_Click
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.Save
Command7.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command10_Click
End Sub

Private Sub Command8_Click()
If Adodc1.Recordset.RecordCount = 1 Then
    Adodc1.Recordset.Delete
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
ElseIf Adodc1.Recordset.RecordCount = 0 Then
    MsgBox ("No Data in the DataBase.")
Else
    Adodc1.Recordset.Delete
    Command3_Click
End If
Command10_Click
End Sub

Private Sub Command9_Click()
Command7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Command10_Click
End Sub
Private Sub Form_Load()
    Command10_Click
End Sub
