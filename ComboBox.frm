VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   LinkTopic       =   "Form14"
   ScaleHeight     =   1995
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4680
      TabIndex        =   2
      Text            =   "Font Size"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Text            =   "Font Type"
      Top             =   1200
      Width           =   2055
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
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Text            =   "STOP STARING AT MY SCREEN"
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Text1.Font = Combo1.Text
End Sub

Private Sub Combo2_Click()
Text1.FontSize = Combo2.Text
End Sub

Private Sub Form_Load()
With Combo1
    .AddItem "Arial"
    .AddItem "Calibri"
    .AddItem "Impact"
    .AddItem "Comic Sans MS"
    .AddItem "Georgia"
    .AddItem "Modern"
End With
With Combo2
    .AddItem "8"
    .AddItem "10"
    .AddItem "14"
    .AddItem "18"
    .AddItem "20"
    .AddItem "24"
End With
End Sub
