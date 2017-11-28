VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12780
   LinkTopic       =   "Form5"
   ScaleHeight     =   7110
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   975
      Left            =   4920
      TabIndex        =   16
      Top             =   4800
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GRADE"
      Height          =   2175
      Left            =   8640
      TabIndex        =   14
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Height          =   975
      Left            =   4920
      TabIndex        =   13
      Top             =   3600
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULATE AVERAGE"
      Height          =   2175
      Left            =   4680
      TabIndex        =   6
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label7 
      Caption         =   "GRADE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   15
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "AVERAGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   7
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "SUBJECT 5"
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "SUBJECT 4"
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "SUBJECT 3"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "SUBJECT 2"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "SUBJECT 1"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim avg As Long
Dim s1 As Long
Dim s2 As Long
Dim s3 As Long
Dim s5 As Long
Dim s4 As Long
Private Sub Command1_Click()
Dim flag As Integer
flag = 1
s1 = Val(Text1.Text)
s2 = Val(Text2.Text)
s3 = Val(Text3.Text)
s4 = Val(Text4.Text)
s5 = Val(Text5.Text)
avg = (s1 + s2 + s3 + s4 + s5) / 5
If s1 > 100 Or s2 > 100 Or s3 > 100 Or s4 > 100 Or s5 > 100 Then
MsgBox ("Error in input")
flag = 0
ElseIf s1 < 0 Or s2 < 0 Or s3 < 0 Or s4 < 0 Or s5 < 0 Then
MsgBox ("Error in input")
flag = 0
End If
If flag = 1 Then
Text6.Text = avg
End If
End Sub

Private Sub Command2_Click()
If avg > 90 Then
Text7.Text = "O"
ElseIf avg <= 90 And avg > 80 Then
Text7.Text = "E"
ElseIf avg <= 80 And avg > 70 Then
Text7.Text = "A"
ElseIf avg <= 70 And avg > 60 Then
Text7.Text = "B"
ElseIf avg <= 60 And avg > 50 Then
Text7.Text = "C"
ElseIf avg <= 50 And avg > 40 Then
Text7.Text = "D"
ElseIf avg <= 40 Then
Text7.Text = "F"
End If
End Sub

