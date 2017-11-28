VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "Loan Calculator"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3825
   LinkTopic       =   "Form15"
   ScaleHeight     =   2985
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Payment"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check if early payment"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Installment"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Duration"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Interest Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Loan Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Long
Dim r As Double
Dim t As Double
Dim i As Long
Dim s As Integer

Private Sub Command1_Click()
i = Pmt((r / (12 * 100)), t, -p, 0, Check1.Value)
Text4.Text = i
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Text1_Change()
p = Val(Text1.Text)
End Sub

Private Sub Text2_Change()
r = Val(Text2.Text)
End Sub

Private Sub Text3_Change()
t = Val(Text3.Text)
End Sub
