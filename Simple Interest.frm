VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "SIMPLE INTEREST"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Double, r As Double, t As Double, si As Double

Private Sub Form_Load()
p = Val(InputBox("Enter the Principal amount"))
r = Val(InputBox("Enter the Rate of interest"))
t = Val(InputBox("Enter the Time"))
si = (p * r * t) / 100
Text1.Text = si
End Sub

