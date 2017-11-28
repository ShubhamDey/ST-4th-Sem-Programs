VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Form13"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   LinkTopic       =   "Form13"
   ScaleHeight     =   3810
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Back Color"
      Height          =   2295
      Left            =   5760
      TabIndex        =   11
      Top             =   1200
      Width           =   2535
      Begin VB.OptionButton Option8 
         Caption         =   "Red"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Green"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Blue"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         Caption         =   "None"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   2055
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   240
      TabIndex        =   2
      Text            =   "STOP STARING AT MY SCREEN"
      Top             =   240
      Width           =   8055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Font"
      Height          =   2295
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
      Begin VB.CheckBox Check4 
         Caption         =   "Strike-Through"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Underline"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italic"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fore Color"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
      Begin VB.OptionButton Option4 
         Caption         =   "None"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Blue"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Green"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Red"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Text1.FontBold = True Then
    Text1.FontBold = False
Else
Text1.FontBold = True
End If
End Sub

Private Sub Check2_Click()
If Text1.FontItalic = True Then
    Text1.FontItalic = False
Else
Text1.FontItalic = True
End If
End Sub

Private Sub Check3_Click()
If Text1.FontUnderline = True Then
    Text1.FontUnderline = False
Else
Text1.FontUnderline = True
End If
End Sub

Private Sub Check4_Click()
If Text1.FontStrikethru = True Then
    Text1.FontStrikethru = False
Else
Text1.FontStrikethru = True
End If
End Sub

Private Sub Option1_Click()
Text1.ForeColor = vbRed
End Sub

Private Sub Option2_Click()
Text1.ForeColor = vbGreen
End Sub

Private Sub Option3_Click()
Text1.ForeColor = vbBlue
End Sub

Private Sub Option4_Click()
Text1.ForeColor = vbnone
End Sub

Private Sub Option5_Click()
Text1.BackColor = vbWhite
End Sub

Private Sub Option6_Click()
Text1.BackColor = vbBlue
End Sub

Private Sub Option7_Click()
Text1.BackColor = vbGreen
End Sub

Private Sub Option8_Click()
Text1.BackColor = vbRed
End Sub
