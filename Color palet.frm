VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll3 
      Height          =   495
      Left            =   1920
      Max             =   255
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   495
      Left            =   1920
      Max             =   255
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   1920
      Max             =   255
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HScroll1_Change()
Label4.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Label1.BackColor = RGB(HScroll1.Value, 0, 0)
End Sub

Private Sub HScroll2_Change()
Label4.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Label2.BackColor = RGB(0, HScroll1.Value, 0)
End Sub

Private Sub HScroll3_Change()
Label4.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Label3.BackColor = RGB(0, 0, HScroll1.Value)
End Sub
