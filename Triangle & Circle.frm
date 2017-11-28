VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   LinkTopic       =   "Form6"
   ScaleHeight     =   6270
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "AREA OF TRIANGLE"
      Height          =   3015
      Left            =   5160
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AREA OF CIRCLE"
      Height          =   3015
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "AREA OF TRIANGLE AND CIRCLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   7695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim areac As Long
Dim areat As Long
Dim base As Long
Dim rad As Long
Dim hgt As Long

Private Sub Command1_Click()
Dim flag As Integer
flag = 1
hgt = Val(InputBox("Enter Height"))
base = Val(InputBox("Enter Radius"))
If hgt < 0 Or base < 0 Then
MsgBox ("Parameters can't be negative")
flag = 0
End If
areat = (base * hgt) / 2
If flag = 1 Then
MsgBox ("The area of the triangle is:" & areat)
End If
End Sub

Private Sub Command2_Click()
Dim flag As Integer
flag = 1
rad = Val(InputBox("Enter radius:"))
If rad < 0 Then
MsgBox ("Radius can't be negative")
flag = 0
End If
areac = 3.14 * rad * rad
If flag = 1 Then
MsgBox ("The area of the circle is: " & areac)
End If
End Sub
