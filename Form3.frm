VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form3"
   ScaleHeight     =   5175
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULATE"
      Height          =   1215
      Left            =   3720
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "FACTORIAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Result"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Input a Number"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer, p As Integer, i As Integer
Private Sub Command1_Click()
n = Val(Text1.Text)
p = 1
For i = 1 To n
p = p * i
Next i
Label2.Caption = p
End Sub

