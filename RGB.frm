VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "DIV"
      Height          =   975
      Left            =   7920
      TabIndex        =   12
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MULTIPLY"
      Height          =   975
      Left            =   5640
      TabIndex        =   11
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SUBSTRACT"
      Height          =   1095
      Left            =   7920
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ADD"
      Height          =   1095
      Left            =   5640
      TabIndex        =   9
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BLUE"
      Height          =   1335
      Left            =   6840
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GREEN"
      Height          =   1335
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RED"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "RESULT"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "2nd Number"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "1st Number"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n1 As Integer, n2 As Integer, n3 As Integer
Private Sub Command1_Click()
Form1.BackColor = vbRed
End Sub

Private Sub Command2_Click()
Form1.BackColor = vbGreen
End Sub

Private Sub Command3_Click()
Form1.BackColor = vbBlue
End Sub

Private Sub Command4_Click()
n1 = Val(Text1.Text)
n2 = Val(Text2.Text)
n3 = n1 + n2
Text3.Text = n3
End Sub

Private Sub Command5_Click()
n1 = Val(Text1.Text)
n2 = Val(Text2.Text)
n3 = n1 - n2
Text3.Text = n3
End Sub

Private Sub Command6_Click()
n1 = Val(Text1.Text)
n2 = Val(Text2.Text)
n3 = n1 * n2
Text3.Text = n3
End Sub

Private Sub Command7_Click()
n1 = Val(Text1.Text)
n2 = Val(Text2.Text)
n3 = n1 / n2
Text3.Text = n3
End Sub
